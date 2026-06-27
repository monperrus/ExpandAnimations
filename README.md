ExpandAnimations is a LibreOffice/OpenOffice.org extension to expand animations into static slides before exporting to PDF. The extension adds a menu entry "Tools>>Add-ons>>Expand animations" in Impress. The generated PDF and editable expanded ODP files are saved in the same folder as the source document. For a presentation named `presentation.odp`, the extension creates `presentation.pdf` and an animation-free `presentation-expanded.odp`. Hidden slides are omitted from the expanded ODP so re-exporting it keeps the same page order and internal links as the generated PDF.

**We are looking for contributors, knowledgeable in Basic, see [issue list](https://github.com/monperrus/ExpandAnimations/issues?q=is%3Aissue+is%3Aopen+sort%3Aupdated-desc) :)** 

Usage
-------

Download: see latest file at <https://github.com/monperrus/ExpandAnimations/releases> (open the OXT file with LibreOffice)

Command line usage
------------------

The extension can be run from the command line after it has been installed in LibreOffice:

```bash
EXPANDANIMATIONS_INPUT=/path/to/presentation.odp libreoffice --headless "macro:///ExpandAnimations.ExpandAnimations.CommandLine"
```

The `CommandLine` macro expands the presentation without showing a completion dialog, creates the same output files as the menu entry (`presentation-expanded.odp` and `presentation.pdf`), then exits LibreOffice.

The same macro can be called through `soffice` or `soffice.bin` when those executables are available in your LibreOffice installation:

```bash
EXPANDANIMATIONS_INPUT=/path/to/presentation.odp soffice --headless "macro:///ExpandAnimations.ExpandAnimations.CommandLine"
```

On Linux desktops such as GNOME/Nautilus, this can be used from a right-click script. Save a script like this as `~/.local/share/nautilus/scripts/ODP-to-expanded-PDF.sh`, make it executable, then run it from the file manager scripts menu:

```bash
#!/bin/bash
set -e

input_file="$(realpath "$1")"
EXPANDANIMATIONS_INPUT="$input_file" libreoffice --headless "macro:///ExpandAnimations.ExpandAnimations.CommandLine"
```

Current limitations
-------------------

ExpandAnimations currently expands visibility-based animation steps, such as appear and disappear effects. Animated GIF playback cannot be preserved in the generated PDF because PDF viewers do not support GIF animation natively. Other animation families, such as transparency changes, fill-color changes, font-color changes, or other emphasis effects, are not expanded into intermediate static slides yet. Slides that contain only unsupported animation effects are kept as single static slides; the interactive menu command reports a warning when this happens.

Issue tracker: <https://github.com/monperrus/ExpandAnimations/issues>

License
--------

Copyright (c) 2011  Matthew Neeley, Martin Monperrus
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Lesser General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

It is a fork of an excellent macro written by Matthew Neeley (see <http://markmail.org/message/ewe336yoe6iennsf> and <https://gist.github.com/977752>).
The extension code was initially generated using the excellent BasicAddonBuilder (<http://extensions.services.openoffice.org/en/project/BasicAddonBuilder>),
