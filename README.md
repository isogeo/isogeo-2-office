[![Build Status](https://travis-ci.org/isogeo/isogeo-2-office.svg?branch=master)](https://travis-ci.org/isogeo/isogeo-2-office)

Isogeo to Office
================

Desktop toolbox using Isogeo REST API to export metadatas into Microsoft Office files and standardized XML (ISO 19139).

Funded by:

![Agence des Espaces Verts d'Île de France](/img/logo_aev.jpg) ![Conseil Départemental du Loiret](/img/logo_cd45.jpg)

**Usage**:

* [French : documentation utilisateur](https://www.gitbook.com/book/isogeo/app-isogeo2office/details) ;
* [English: user documentation](https://www.gitbook.com/book/isogeo/app-isogeo2office/details).

## Tips

### Shortcut

Create a Windows shortcut: Right clic > New > Shortcut and insert this command replacing with the absolute paths (removing brackets): `C:\Windows\System32\cmd.exe /k "{absolute_path_to_the_folder}\isogeo2office\virtenv\Scripts\python {absolute_path_to_the_folder}\isogeo2office\isogeo2office.py"`

## Scheduled task

Program/script:

`{absolute_path_to_the_folder}\virtenv\Scripts\python.exe`

Arguments:

`{absolute_path_to_the_folder}\isogeo2office.py 0`

Launch in :

`{absolute_path_to_the_folder}\`

## Support

This application is not part of Isogeo license contract and won't be supported or maintained as well. If you need help, send a mail to <projets+isogeo2office@isogeo.fr>
