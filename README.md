Phlegmoirs: a text editor written in visual basic intended for simple journaling, and very quick browsing and editing of text files.

FEATURES:

* TODO: Write a features list.
* TODO: Explain how to install the project, and how to get VB6 working smoothly in 2023.

ICONS:

As of v0.22.0, most icons were scrapped and replaced by selections from the [Fugue Icons Pack](https://www.iconarchive.com/show/fugue-icons-by-yusuke-kamiyamane.html) by [Yusuke Kamiyamane](https://p.yusukekamiyamane.com), used under the [Creative Commons Attribution 4.0 License](https://creativecommons.org/licenses/by/4.0/). When it's not Fugue, it's the [Farm Fresh Icons Pack](https://www.iconarchive.com/show/farm-fresh-icons-by-fatcow.html) by [Fatcow Web Hosting](https://fatcow.com/free-icons) under the same license. In older versions I was using system icons found scattered around my Windows 2000 installation, and some I made myself. And sometimes I just put a letter on the button, like "R" for Replace, and called it a day. (Even though I had already used "R" for Refresh, I called it, it's a day.)

I am required to state that some Fugue ones were modified. The split file browser icon didn't come in blue, so I spliced in the grey split screen. Folder-Up is an Up pasted onto a Folder. The blue floppy disk icon wasn't big enough, so I cut it in half, on both axes, and pasted some of its cross-section into the extra space. Is that noticeable? I feel like everybody who sees it will immediately say "that's not a proportional drawing of a floppy disk... did somebody TAKE a proportional drawing of a floppy disk and RUIN it?! Sure looks that way!"

EXTERNAL TOOLS:

Modifying those icons in a way that looked decent in VB6 would have been futile without the free [Junior Icon Editor](http://www.sibcode.com/junior-icon-editor/index.htm). Then to make an icon into a cursor with hotspot coordinates takes the shareware [SIB Cursor Editor](http://www.sibcode.com/cursor-editor/). These folks make an [icon extractor](http://www.sibcode.com/icon-extractor/index.htm), too, in case you thought supplying the VB editor with an image loses it forever.

Also, in a well-intended effort at turd-polishing, we are doing some refactors to acknowledge that coding practices are a thing these days.

[MZ-Tools 8 for VB 6.0](https://www.mztools.com) is being helpful toward this end, and some modified configuration files are included in the tools folder here so as to stay on the same page, standards-wise. They would be inserted at roughly:

D:\AppData\Roaming\MZTools Software\MZTools8\VB6\TeamOptions

Beyond code quality analysis, it adds navigation features to the Visual Basic IDE.

[X-Mouse Button Control](https://www.highrez.co.uk/downloads/xmousebuttoncontrol.htm) makes the mousewheel function in the IDE. Shared config files seem to not transfer well, but setup is trivial anyway. Profiles can be broken down to the window level so that it's not affecting, for instance, an editor window that needed fixing while also an MZ-Tools window that did not need fixing.

With some setup, VB6 is surprisingly usable in 2023 without having to drop the IDE for another code editor. It only requires *not stopping to wonder if you should*.
