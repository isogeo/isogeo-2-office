QSettings
=========

Application and user settings are managed using [QSettings](http://doc.qt.io/qt-5/qsettings.html).

According to [the official documentation](http://doc.qt.io/qt-5/qsettings.html#locations-where-application-settings-are-stored), on Windows, settings are stored directly in the registry (depending on the user who launches the application):

1. `HKEY_CURRENT_USER\Software\MySoft\isogeoToOffice`
2. `HKEY_LOCAL_MACHINE\Software\MySoft\isogeoToOffice`

For example, on Windows 10:

![](https://raw.githubusercontent.com/isogeo/isogeo-2-office/master/img/docs/settings_win_registry.png)

