import { useState } from "react";

export function usePreviewAuthorizationVersion(
  requestKey: string,
  refresh: () => Promise<string | null>,
) {
  const [request, setRequest] = useState({ requestKey, refresh, version: 0 });
  if (request.requestKey !== requestKey || request.refresh !== refresh) {
    const version = request.version + 1;
    setRequest({ requestKey, refresh, version });
    return version;
  }
  return request.version;
}
