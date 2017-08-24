
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

*Q: What to do?*

**A: First**,
  use the built-in "Compress Pictures" feature: Go "File ->
  Compress pictures", or select any image, go to the "Picture Tools"
  toolbar, and select the "Compress Pictures" icon (four arrows pointing
  to the corners of an image). This tool allows you to down-scale pictures
  and removed cropped-out areas, and can be applied to all pictures in the
  presentation at once, *but does not change the image format of pictures
  in the presentation.*
  Make sure to save your presentation under a new name, in case you realize
  you need to some of the original, uncompressed pictures!

**A: Then**,
  if the presentation file size is still excessive, use ``pptx-downsizer``!

``pptx-downsizer`` will go over all images in your presentation (pptx),
and down-size all images above a certain size.

-  By default, all images are converted to PNG format (except for JPEGs
   which remains in JPEG format). This is particularly relevant if you
   have a lot of TIFF files in your presentation, e.g. if you copy/paste
   images or use the Mac "screen capture" feature when adding images.
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
remaining large/complex PNG images to JPEG, as a separate downsizing step,
you should be able to get another 20-40 percent reduction. Of course, this
all depends on how large and complex your original images are, and how
much you are willing to compromise quality when compressing your images.
You can use the ``--quality`` parameter to adjust quality of JPEG images.


How ``pptx-downsizer`` works:
-----------------------------

1. First it unzips the ``.pptx`` PowerPoint file to a temporary directory.
   Other ooxml files probably works as well, e.g. ``.docx`` Word files.
2. Then, ``pptx-downsizer`` searches for image files with large file size.
   The file size is controlled with the ``fsize-filter`` parameter.
   It is possible to add additional file-selection filter criteria,
   e.g. set ``fname-filter="*.TIFF"`` to only convert TIFF image files,
   although this is typically not needed.
3. ``pptx-downsizer`` will then go through all selected images and try
   to minimize them in the following ways:

   a. If the image dimensions are larger than ``img-max-size``,
      ``pptx-downsizer`` will reduce the image dimensions (by an interger
      factor) so that the image is smaller than ``img-max-size``.
   b. The image is then resaved in the selected format (default: jpeg) and
      quality (default: 90). ``pptx-downsizer`` can also be used to change
      image modes, e.g. convert transparent regions of PNG images to a solid
      color by setting ``--img-mode="rgb" --fill-color="#ffffff"``.

4. Finally, the ``.pptx`` PowerPoint file is re-created and re-saved as
   ``Presentation.downsized.pptx``.

Note: It is often useful to do multiple rounds of downsizing, e.g. first
converting all large TIFF files to PNG format, then downsizing the downsized
``pptx`` to convert the biggest PNG images to JPEG (see "Examples" below).



Examples usage:
---------------

Make sure to save your presentation (and, preferably exit PowerPoint,
and make a backup of your presentation just in case).

Let's say you have your original, large presentation saved as
``Presentation.pptx``

After installing ``pptx-downsizer``, you can run the following from your
terminal::

    pptx-downsizer "Presentation.pptx"

If you want to change the file size limit used to determine what images
are down-sized to 1 MB (≈ 1'000'000 bytes)::

    pptx-downsizer "Presentation.pptx" --fsize-filter 1e6

If you want to disable down-scaling of large high-resolution images, set
``img-max-size`` to 0::

    pptx-downsizer "Presentation.pptx" --img-max-size 0

If you want to convert large images to JPEG format::

    pptx-downsizer "Presentation.pptx" --convert-to jpeg

**Advanced usage:** Pause before re-creating the PowerPoint file.
Let's say you are a power user, and you need to do something very specific
to some or all of the images in your presentation. For instance, adding
watermarks before sending the presentation to someone else.
If you pass ``--wait-before-zip`` to ``pptx-downsizer``, the program will
wait before it re-creates the presentation (but after downsizing the images).




Command line arguments:
-----------------------

You can always get a complete description of the program and the
available command line arguments (parameters) by invoking::

    pptx-downsizer --help


This should produce an output similar to the following::

    $ pptx-downsizer --help
    usage: pptx-downsizer [-h] [--fname-filter GLOB] [--fsize-filter SIZE]
                          [--convert-to IMAGE_FORMAT] [--img-max-size PIXELS]
                          [--img-mode MODE] [--fill-color COLOR]
                          [--quality [1-100]] [--optimize] [--no-optimize]
                          [--outputfn_fmt FORMAT-STRING] [--overwrite]
                          [--compress-type ZIP-TYPE] [--wait-before-zip]
                          [--on-error DO-WHAT] [--verbose [0-5]]
                          filename

    PowerPoint pptx downsizer. Reduce the file size of PowerPoint presentations by
    re-compressing images within the pptx file.

    positional arguments:
      filename              Path to the PowerPoint pptx file that you want to
                            down-size.

    optional arguments:
      -h, --help            show this help message and exit
      --fname-filter GLOB   Convert all images matching this filename pattern,
                            e.g. '*.TIFF' (default: None)
      --fsize-filter SIZE   Convert all images with a current file size exceeding
                            this limit, e.g. '1e6' for 1 MB. (default: 524288)
      --convert-to IMAGE_FORMAT
                            Convert images to this image format, e.g. `png` or
                            `jpeg`. (default: png)
      --img-max-size PIXELS
                            If images are larger than this size (width or height),
                            reduce/downscale the image size to make it less than
                            this size. (default: 2048)
      --img-mode MODE       Convert images to this image mode before saving them,
                            e.g. 'RGB' - advanced option. (default: None)
      --fill-color COLOR    If converting image mode (e.g. from RGBA to RGB), use
                            this color for transparent regions. (default: None)
      --quality [1-100]     Quality of converted images (only applies to jpeg
                            output). (default: 90)
      --optimize            Try to optimize the converted image output when
                            saving. Optimizing the output may produce better
                            images, but disabling it may make the conversion run
                            faster. Enabled by default. (default: True)
      --no-optimize         Disable optimization. (default: False)
      --outputfn_fmt FORMAT-STRING
                            How to format the downsized presentation pptx filename
                            Slightly advanced, uses python string formatting.
                            (default: {fnroot}.downsized.pptx)
      --overwrite           Whether to silently overwrite existing file if the
                            output filename already exists. (default: None)
      --compress-type ZIP-TYPE
                            Which zip compression type to use, e.g. ZIP_DEFLATED,
                            ZIP_BZIP2, or ZIP_LZMA. (default: ZIP_DEFLATED)
      --wait-before-zip     If this flag is specified, the program will wait after
                            converting all images before re-zipping the output
                            pptx file. You can use this to make manual changes to
                            the presentation - advanced option. (default: False)
      --on-error DO-WHAT    What to do if the program encounters any errors during
                            execution. `continue` will cause the program to
                            continue even if one or more images fails to be
                            converted. (default: raise)
      --verbose [0-5]       Increase or decrease the 'verbosity' of the program,
                            i.e. how much information it prints about the process.
                            (default: 2)



Installation:
-------------

First, make sure you have Python 3+ installed. I recommend using the
Anaconda Python distribution, which makes everything a lot easier.

With python installed, install ``pptx-downsizer`` using ``pip``::

    pip install pptx-downsizer

You can make sure ``pptx-downsizer`` is installed by calling it
anywhere from the terminal / command prompt::

    pptx-downsizer

Note: You may want to install ``pptx-downsizer`` in a
separate/non-default python environment. If you know what that means,
you already know how to do that. If you do not know what that means,
then don't worry–you probably don't need it after all.


Troubleshooting and bugs:
-------------------------

**NOTE:** ``pptx-downsizer`` is very early/beta software. I strongly
recommend to (a) *back up your presentation to a separate folder before
running* ``pptx-downsizer``, and (b) *work for as long as possible in
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

*OBS: If PowerPoint gives you errors when opening the downsized file,
please don't bother trying to fix the downsized file yourself. You may
run into unexpected errors later. Instead, just continue working with
your original presentation.*

Q: Why doesn't ``pptx-downsizer`` work?

A: It works for me and all the ``.pptx`` files I've thrown at it.
However, there are obviously going to be a lot of scenarios that I
haven't run into yet.

Q: Does ``pptx-downsizer`` overwrite the original presentation file?

A: No, by default ``pptx-downsizer`` will create a new file with
".downsized" added to the filename. If this output file already exists,
``pptx-downsizer`` will let you know, giving you a change to (manually)
move/rename the existing file if you want to keep it. You can disable
this prompt using the ``--overwrite`` argument.
