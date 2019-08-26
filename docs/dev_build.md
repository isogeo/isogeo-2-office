Build and tests
===============

# Tests

Unit tests are runned for each commit on TravisCI.

----

# Build

## Continuous integration

Application is built using AppVeyor ([see configuration in appveyor.yml]()).

## Build locally

### Requirements

First of all, install requirements:

* Windows 10+
* Python 3.7.x

### Visual Studio Code

With Visual Studio Code, using defined tasks:

![](https://raw.githubusercontent.com/isogeo/isogeo-2-office/master/img/docs/build_vsc_tasks.png)

### Manually

Run `build.ps1` script from repository root folder.
