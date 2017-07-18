
pptx-downsizer
==============

Python tool for downsizing Microsoft PowerPoint presentations (pptx)
files.

https://github.com/scholer/pptx-downsizer

Currently only supports downsizing of images (not e.g. videos and other
media files).


Use cases:
----------

Why might someone want to downsize a Power Point presentation?

If you are like me, when you create a PowerPoint presentation, you just
throw in a lot of images on the slides without paying too much attention
to how large the images are.

You may even use the "screenshot" feature (Cmd+shift+4 on Mac) to
quickly capture images of whatever you have on your screen, and paste it
into the PowerPoint presentation (using "Paste special"). In which case
you are actually creating large TIFF images in your presentation (at
least for PowerPoint 2016).

Even though the images in the presentation are compressed/zipped when
saving the presentation file, the presentation will still be
significantly larger than it actually needs to be.

However, once you realize that your presentation is 100+ MBs, you don't
have the time to re-save a lower-quality version of each image and then
substitute that image in the presentation.

Q: What to do?

A: **First**, use the built-in "Compress Pictures" feature: Go "File ->
Compress pictures", or select any image, go to the "Picture Tools"
toolbar, and select the "Compress Pictures" icon (four arrows pointing
to the corners of an image). This tool allows you to down-scale pictures
and removed cropped-out areas, and can be applied to all pictures in the
presentation at once, *but does not change the image format of pictures
in the presentation.*

Make sure to save your presentation under a new name, in case you want
to revert some of the compressed pictures!

A: **Then**, use ``pptx-downsizer``!

``pptx-downsizer`` will go over all images in your presentation (pptx),
and down-size all images above a certain size.

-  By default, all images are converted to PNG format (except for JPEGs
   which remains in JPEG format).
-  You can also choose to use JPEG format (recommended only after doing
   an initial downsizing using PNG).
-  If images are more than a certain limit (default 2048 pixels) in
   either dimension (width, height), they are down-scaled to a more
   reasonable size (you most likely do not need very high-resolution
   images in your presentation, since most projectors still have a
   relatively low resolution anyways.)

Q: How much can I expect ``pptx-downsizer`` to reduce my powerpoint
presentations (pptx files)?

A: If you have copy/pasted a lot of screenshots (TIFF files), it is not
uncommon to for the presentation to be reduced to less than half (and in
some case one fourth) of the original file size. If you further convert
remaining large PNG images to JPEG (as a separate downsizing), you
should be able to get another 30-50 percent reduction. Of course, this
all depends on how large your original images are and how much you are
willing to


Examples:
---------

Make sure to save your presentation (and, preferably exit PowerPoint,
and make a backup of your presentation just in case).

Let's say you have your original, large presentation saved as
``Presentation.pptx``

After installing ``pptx-downsizer``, you can run the following from your
terminal (notice the substitution of the hyphen for an underscore)::

    pptx_downsizer "Presentation.pptx"

If you want to change the file size limit used to determine what images
are down-sized to 1 MB (≈ 1'000'000 bytes)::

    pptx_downsizer "Presentation.pptx" --fsize-filter 1e6

If you want to disable down-scaling of large high-resolution images, set
``img-max-size`` to 0::

    pptx_downsizer "Presentation.pptx" --img-max-size 0

If you want to convert large images to JPEG format::

    pptx_downsizer "Presentation.pptx" --convert-to jpeg


Installation:
-------------

First, make sure you have Python 3+ installed. I recommend using the
Anaconda Python distribution, which makes everything a lot easier.

With python installed, install ``pptx-downsizer`` using ``pip``::

    pip install pptx-downsizer

You can make sure ``pptx-downsizer`` is installed by invoking it
anywhere from the terminal (command line)::

    pptx_downsizer

Note: You may want to install ``pptx-downsizer`` in a
separate/non-default python environment. (If you know what that means,
you already know how to do that. If you do not know what that means,
then don't worry–you probably don't need it after all).


Troubleshooting and bugs:
-------------------------

**NOTE:** ``pptx-downsizer`` is very early/beta software. I strongly
recommend to (a) *back up your presentation to a separate folder before
running* ``pptx_downsizer``, and (b) *work for as long as possible in
the original presentation.* That way, if ``pptx-downsizer`` doesn't
work, you can always go back to your original presentation, and you will
not have lost any work.

Q: HELP! I ran the downsizer and now the presentation won't open or
PowerPoint gives errors when opening the pptx file!

A: Sorry that ``pptx-downsizer`` didn't work for you. If you want, feel
free to send me a copy of both the presentation and the downsized pptx
file produced by this script, and I'll try to figure out what the
problem is. There are, unfortunately, a lot of things that could be
wrong, and without the original presentation, I probably cannot diagnose
the issue.

**OBS: If PowerPoint gives you errors when opening the downsized file,
please don't bother trying to fix the downsized file yourself. You may
run into unexpected errors later. Instead, just continue working with
your original presentation.**

Q: Why doesn't ``pptx-downsizer`` work?

A: It works for me and all the ``.pptx`` files I've thrown at it.
However, there are obviously going to be a lot of scenarios that I
haven't run into yet.

Q: Does ``pptx-downsizer`` overwrite the original presentation file?

A: No, by default ``pptx-downsizer`` will create a new file with
".downsized" postfix. If the intended output already exists,
``pptx_downloader`` will let you know, giving you a change to (manually)
move/rename the existing file if you want to keep it. You can disable
this prompt using the ``--overwrite`` argument.
