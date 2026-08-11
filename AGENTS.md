# Repository instructions

- Read `README.md`, `docs/HANDOVER.md`, and `docs/UPDATE_GUIDE.md` before changing release, update, rollback, or user-data behavior.
- Keep `app/app_main.py` and `app/app_main.pyw` synchronized. Do not commit credentials, DPAPI data, browser profiles, user configuration, publish state, logs, screenshots, generated installers, or update ZIP files.
- Preserve `%LOCALAPPDATA%\DouyinPublisher` data and user-selected images, spreadsheets, installation paths, and progress across install and update operations.
- Run the closest syntax, self-check, core regression, package-content, SHA-256, and installed-startup checks applicable to the change before declaring it complete.

## Rollback retention and disk hygiene

- On local disks, keep the current verified installer/update package plus the newest verified previous stable rollback generation. After a newer rollback generation passes signature/hash, package-content, install/startup, and user-data preservation checks, remove older rebuildable local installers, update ZIP duplicates, extracted validation directories, temporary build outputs, and caches for that same release scope.
- Do not automatically delete GitHub Releases or any asset referenced by `version.json`, the unified management platform, an active rollout, or the newest rollback instructions. Never remove user data, secrets, the newest verified rollback, active signed artifacts, or Git history.
- Resolve and verify every absolute cleanup target before deletion, ensure it is a generated rollback/build location rather than a user-data path, and report removed paths and recoverability. If the newest rollback or target scope is uncertain, do not delete it.
