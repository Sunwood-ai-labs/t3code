// @effect-diagnostics nodeBuiltinImport:off - a native reader process verifies Windows file sharing.
import * as NodeChildProcess from "node:child_process";
import * as NodeServices from "@effect/platform-node/NodeServices";
import { assert, it } from "@effect/vitest";
import * as HostProcess from "@t3tools/shared/HostProcess";
import * as NodeSqliteClient from "@t3tools/shared/nodeSqliteClient";
import * as Effect from "effect/Effect";
import * as FileSystem from "effect/FileSystem";
import * as Path from "effect/Path";
import * as PlatformError from "effect/PlatformError";
import * as SqlClient from "effect/sql/SqlClient";
import * as Scope from "effect/Scope";

import * as DatabaseMaintenance from "./DatabaseMaintenance.ts";

const readerSource = `
const { DatabaseSync } = require("node:sqlite");
const db = new DatabaseSync(process.argv[1]);
db.exec("PRAGMA journal_mode = WAL");
const rows = db.prepare("SELECT value FROM retained_data").all();
process.stdout.write(JSON.stringify(rows) + "\\n");
process.stdin.resume();
process.stdin.once("end", () => db.close());
`;

const openReader = (executable: string, filename: string) =>
  Effect.acquireRelease(
    Effect.sync(() => NodeChildProcess.spawn(executable, ["-e", readerSource, filename])),
    (child) =>
      Effect.promise(() => {
        if (child.pid === undefined || child.exitCode !== null || child.signalCode !== null)
          return Promise.resolve();
        return new Promise<void>((resolve) => {
          child.once("exit", () => resolve());
          child.stdin.end();
        });
      }),
  ).pipe(
    Effect.tap((child) =>
      Effect.promise(
        () =>
          new Promise<void>((resolve, reject) => {
            child.stdout.once("data", (data) => {
              const ready = String(data).trim();
              if (ready === '[{"value":"keep"}]') resolve();
              else reject(new Error(`Unexpected SQLite reader response: ${ready}`));
            });
            child.once("error", reject);
            child.once("exit", () => reject(new Error("SQLite reader exited before opening")));
          }),
      ),
    ),
  );

it.layer(NodeServices.layer)("legacy database removal boundaries", (it) => {
  it.effect.skipIf(HostProcess.Platform.defaultValue() !== "win32")(
    "preserves the database and sidecars if a foreign reader opens after the lock check",
    () =>
      Effect.gen(function* () {
        const fs = yield* FileSystem.FileSystem;
        const path = yield* Path.Path;
        const executable = yield* HostProcess.ExecutablePath;
        const home = yield* fs.makeTempDirectoryScoped({ prefix: "t3-legacy-handoff-" });
        const databasePath = path.join(home, "statev2.sqlite");
        const legacy = path.join(home, "state.sqlite");
        yield* fs.writeFileString(databasePath, "current database retained");
        yield* Effect.gen(function* () {
          const sql = yield* SqlClient.SqlClient;
          yield* sql`PRAGMA journal_mode = WAL`;
          yield* sql`CREATE TABLE retained_data (value TEXT)`;
          yield* sql`INSERT INTO retained_data VALUES ('keep')`;
        }).pipe(Effect.provide(NodeSqliteClient.layer({ filename: legacy })));
        const removals: string[] = [];
        yield* Effect.gen(function* () {
          const scope = yield* Scope.Scope;
          const guardedFs = FileSystem.FileSystem.of({
            ...fs,
            remove: (filename, options) =>
              Effect.gen(function* () {
                removals.push(filename);
                if (filename === legacy)
                  yield* openReader(executable, legacy).pipe(
                    Effect.provideService(Scope.Scope, scope),
                  );
                return yield* fs.remove(filename, options);
              }),
          });
          const maintenance = yield* DatabaseMaintenance.DatabaseMaintenance.pipe(
            Effect.provide(DatabaseMaintenance.layer),
            Effect.provideService(FileSystem.FileSystem, guardedFs),
          );
          const failure = yield* Effect.flip(maintenance.removeLegacyDatabase(databasePath));
          assert.equal(failure._tag, "DatabaseInUseError");
          assert.deepEqual(removals, [legacy]);
          assert.isTrue(yield* fs.exists(legacy));
          assert.isTrue(yield* fs.exists(`${legacy}-wal`));
          assert.isTrue(yield* fs.exists(`${legacy}-shm`));
          const rows = yield* Effect.gen(function* () {
            const sql = yield* SqlClient.SqlClient;
            return yield* sql<{ value: string }>`SELECT value FROM retained_data`;
          }).pipe(Effect.provide(NodeSqliteClient.layer({ filename: legacy, readonly: true })));
          assert.deepEqual([...rows], [{ value: "keep" }]);
          assert.equal(yield* fs.readFileString(databasePath), "current database retained");
        }).pipe(Effect.scoped);
        const maintenance = yield* DatabaseMaintenance.DatabaseMaintenance.pipe(
          Effect.provide(DatabaseMaintenance.layer),
        );
        yield* maintenance.removeLegacyDatabase(databasePath);
        assert.isFalse(yield* fs.exists(legacy));
      }).pipe(Effect.scoped),
  );

  it.effect("reports permission failures as removal errors instead of an active database", () =>
    Effect.gen(function* () {
      const fs = yield* FileSystem.FileSystem;
      const path = yield* Path.Path;
      const home = yield* fs.makeTempDirectoryScoped({ prefix: "t3-legacy-permission-" });
      const legacy = path.join(home, "state.sqlite");
      yield* Effect.gen(function* () {
        const sql = yield* SqlClient.SqlClient;
        yield* sql`CREATE TABLE retained_data (value TEXT)`;
      }).pipe(Effect.provide(NodeSqliteClient.layer({ filename: legacy })));
      const denied = PlatformError.systemError({
        _tag: "PermissionDenied",
        module: "FileSystem",
        method: "remove",
        pathOrDescriptor: legacy,
      });
      const maintenance = yield* DatabaseMaintenance.DatabaseMaintenance.pipe(
        Effect.provide(DatabaseMaintenance.layer),
        Effect.provideService(
          FileSystem.FileSystem,
          FileSystem.FileSystem.of({
            ...fs,
            remove: () => Effect.fail(denied),
          }),
        ),
      );
      const failure = yield* Effect.flip(
        maintenance.removeLegacyDatabase(path.join(home, "statev2.sqlite")),
      );
      assert.equal(failure._tag, "LegacyDatabaseRemoveError");
      assert.isTrue(yield* fs.exists(legacy));
    }).pipe(Effect.scoped),
  );
});
