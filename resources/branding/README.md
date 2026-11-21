# FINDEX application and report identity

- favicon.png is the supplied primary desktop artwork; original supplied assets are unchanged.
- icon.ico is derived from that artwork with preserved transparency and aspect ratio.
  It contains multiple Windows icon sizes for the app and installer.
- version_info.txt and installer_version_info.txt provide FINDEX version 3.0.1 identity.
- report_logo.png is the dedicated self-contained HTML report logo asset.

Desktop title/taskbar icons use the included ICO. Sidebar and About image sizes
are unchanged. Sidebar/report logo tiles use the same dark #111625 backing in
both themes. The secondary icon.png is not used as the primary desktop logo.

The build includes this branding directory and resolves it through APP_ROOT in
source and packaged applications. Normal builds do not need to regenerate the ICO.