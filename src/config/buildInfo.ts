export const appVersion = __APP_VERSION__;
export const buildCommit = __BUILD_COMMIT__;
export const buildTimestamp = __BUILD_TIMESTAMP__;

export const appBuildLabel =
  buildCommit && buildCommit !== "local"
    ? `Version ${appVersion} · ${buildCommit}`
    : `Version ${appVersion}`;
