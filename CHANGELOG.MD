The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [Stable Release]

## [2.7.0] - 2017-08-19
### Added
- Ribbon Button to toggle if preview panes are zoomed or not.

## Bugs
- When scrolling fast (holding down the arrow key to scroll) panes don't get zoomed or they get zoomed to much.


## [Stable Release]

## [2.6.1] - 2017-08-15
### Added
- Arrow key scrolling support.

### Fixed
- Fixed issue with tab not reseting causing users to not be able to scroll/loose their tab position.
- Fixed issue with messages loading and zooming before index was set (caused problem with fast comptuers zooming and then unzooming).
- Fix attempt at windows 10 install problems.

## Bugs
- On zoom if a custom tab is select on the ribbon it will loose focus and return to mail items (low priority).
- Windows 10 can't install (problem with test certificate for trusted program).

### Notes
- Totally hacked my way around the non-access API to zoom panels, instead I impersonate a keyboard/mouse zooming until it reaches the zoom level.


## [Beta Release]

## [2.6.0] - 2017-08-15
### Added
- Panel Preview zooming.
- Virtual Keyboard impersonating (used for preview Zoom).
- Mouse movement impersonating (used for preview Zoom).

### Changed
- Improve user friendly errors.
- Adjusted start loop to delay allowing forms to load before zooming (reduces bugs on slow machines).


### Notes
- Totally hacked my way around the non-access API to zoom panels, instead I impersonate a keyboard/mouse zooming until it reaches the zoom level.


## [Stable Release]

## [2.4.0] - 2017-08-10
### Added
- License File.
- Change Log.
- ReadMe.txt file.
- Resource Icon.
- User GUI "Ribbon Interface".
- Office 2013 Support.
- Office 2010 Support.
- Required ZIP Include.
- Todo.txt file.
- VS 2010 Setup Project.

### Changed
- Start using "changelog" over "change log" since it's the common usage.
- Start versioning at v2 (v1 versions considered pre-Alpha).
- Rewrite Ribbon Text.
- Improve exception handling.
- Improve user friendly errors.
- Fix typos exception handling.

### Removed
- Debug Bin.
- Debug Installer.
- Unneeded Registery Entries.

### Notes
- First Release set at Version 2.4.0, download location found at: https://www.robertsroom.com/software/zoomit2010/