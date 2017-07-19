# Copyright 2017, Rasmus S. Sorensen <rasmusscholer@gmail.com>

"""

Module for downsizing Microsoft PowerPoint presentations (pptx) files.

https://github.com/scholer/pptx-downsizer

Currently only supports downsizing of images (not e.g. videos).


"""

# from __future__ import print_function
import argparse
import inspect
import os
import tempfile
import zipfile
from fnmatch import fnmatch
from functools import partial
from glob import glob
import yaml
from PIL import Image

from pptx_downsizer.utils import zip_directory, convert_str_to_int


def downsize_pptx_images(
    filename,
    # Image selection:
    fname_filter=None,  # Only filter files matching this filter (str or callable or None) - or maybe OR filter?
    fsize_filter=int(0.5*2**20),  # Only downsample files above this filesize (number or None)
    img_size_filter=2048,
    # Image conversion/save:
    convert_to="png",
    quality=90,
    optimize=True,
    img_mode=None,
    fill_color=None,  # e.g. '#ffffff',
    # Output pptx file:
    outputfn_fmt="{fnroot}.downsized.pptx",  # "{filename}.downsized.pptx",
    compress_type=zipfile.ZIP_DEFLATED,
    wait_before_zip=False,
    overwrite=None,
    # Program behavior:
    on_error='raise',
    verbose=2,
    # **writer_kwargs
):
    """Downsize a PowerPoint / OfficeOpen pptx file by compressing the images in the presentation.

    Args:
        filename: Filename of the pptx input file.

        fname_filter: Convert images matching this filename glob pattern, e.g. "*.TIFF".
        fsize_filter: Convert images with file size larger than this limit in bytes.
        img_size_filter: Convert images larger than this limit (width or height) in pixels.

        convert_to: Convert images to this image format - e.g. 'png' or 'jpeg'.
        quality: Save images with this quality parameter (JPEG only).
        optimize: Attempt to optimize the image output (for `PIL.Image.save`)
        img_mode: Convert images to this mode before saving - e.g. 'RGB'.
        fill_color: If converting images with alpha channels, use this color as background/fill color.

        outputfn_fmt: The filename format of the generated/downsized pptx file.
        wait_before_zip: If True, prompt the user to press enter before zipping the files in the temporary directory.
        compress_type: Use this zip compression method when making the pptx zip file.
        overwrite: Whether to silently overwrite existing output file if it already exists.

        verbose: Verbosity level, i.e. how much information to print during execution.
        on_error: What to do if the program encounters any error.
            'continue' -> Print error message, then continue.
            'raise'    -> Abort executing and raise error message.

    Returns:
        Filename of the newly generated/downsized pptx file.

    """
    # TODO: If output is .jpg, you may need to add one of the following lines to presentation.xlm:
    #       <Default Extension="jpeg" ContentType="image/jpeg"/>
    #       <Default Extension="jpg" ContentType="application/octet-stream"/>

    # OBS: File endings should be \r\n, even on Mac - because MS software.
    assert os.path.isfile(filename)
    old_fsize = os.path.getsize(filename)
    pptx_fnroot, pptx_ext = os.path.splitext(filename)
    print("\nDownsizing PowerPoint presentation %r (%0.01f MB)...\n" % (filename, old_fsize/2**20))
    convert_to = convert_to.lower().strip(".")
    if convert_to == "jpg":
        print("WARNING: Selected format 'jpg' should be 'jpeg' instead, switching...")
        convert_to = "jpeg"
    if img_mode is None and convert_to == 'jpeg':
        img_mode = 'RGB'
    filter_desc = []
    if fsize_filter:
        filter_desc.append("above %0.01f kB" % (fsize_filter/2**10,))
    if img_size_filter:
        filter_desc.append("larger than %s pixels" % img_size_filter)
    if fname_filter:
        filter_desc.append("with filename matching %r" % fname_filter)

    print(" - Converting image files", " or ".join(filter_desc))
    if isinstance(fname_filter, str):
        fname_filter = partial(fnmatch, pat=fname_filter)

    def ffilter(fname):
        """Return True if file should be included."""
        return (
            (fsize_filter and os.path.getsize(fname) > fsize_filter)
            and (fname_filter is None or fname_filter(fname))
        )

    output_ext = "." + convert_to.strip(".")
    changed_fns = []
    with tempfile.TemporaryDirectory() as tmpdirname:
        pptdir = os.path.join(tmpdirname, "ppt")
        mediadir = os.path.join(pptdir, "media")
        if verbose and verbose > 0:
            print("\nExtracting %r to temporary directory %r..." % (filename, tmpdirname))
        with zipfile.ZipFile(filename, 'r') as zipfd:
            zipfd.extractall(tmpdirname)
        if verbose and verbose > 1:
            print("pptdir:", pptdir)
            print("mediadir:", mediadir)
        image_files = glob(os.path.join(mediadir, "image*"))
        image_files = [fn for fn in image_files if ffilter(fn)]
        print("\nConverting image files...")
        for imgfn in image_files:
            old_img_fsize = os.path.getsize(imgfn)
            print("Converting %r (%s kb)..." % (imgfn,  old_img_fsize//1024))
            fnbase, fnext = os.path.splitext(imgfn)
            if fnext == '.jpg' or fnext == '.jpeg':
                print(" - Preserving JPEG image format for file %r." % (imgfn,))
                outputfn = imgfn
            else:
                outputfn = fnbase + output_ext
            img = Image.open(imgfn)
            if img_size_filter and (img.height > img_size_filter or img.width > img_size_filter):
                downscalefactor = (max(img.size) // img_size_filter) + 1
                newsize = tuple(v // downscalefactor for v in img.size)
                if verbose and verbose > 1:
                    print(" - Resizing %sx, from %s to %s" % (downscalefactor, img.size, newsize))
                img.resize(newsize)
            # extra/unused kwargs to Image.save are silently ignored (e.g. `quality` for png)
            if img_mode:
                if verbose and verbose > 1:
                    print(" - Changing image mode from %s to %s (fill color: %s)..." % (img_mode, img.mode, fill_color))
                if fill_color:
                    # From https://stackoverflow.com/questions/9166400/convert-rgba-png-to-rgb-with-pil
                    img.load()  # needed for split()
                    background = Image.new(img_mode, img.size, fill_color)
                    background.paste(img, mask=img.split()[3])  # 3 is the alpha channel
                    img = background
                else:
                    img = img.convert(img_mode)
            try:
                img.save(outputfn, optimize=optimize, quality=quality)
            except OSError as e:
                if on_error == "continue":
                    print(" - ERROR saving image, skipping!")
                    continue
                else:
                    raise e
            print(" - Saved:  %r (%s kb)" % (outputfn, os.path.getsize(outputfn) // 1024))
            new_img_fsize = os.path.getsize(outputfn)
            if fsize_filter and new_img_fsize > fsize_filter and verbose and verbose > 0:
                print(" - Notice: Filesize %s kb is still above the filesize limit (%s kb)"
                      % (new_img_fsize//1024, fsize_filter//1024))
            if fnext != output_ext:
                # We only need to change the basename, all images are in the same directory...
                changed_fns.append((os.path.basename(imgfn), os.path.basename(outputfn)))
                os.remove(imgfn)
                if verbose and verbose > 1:
                    print(" - Deleted: %r" % (imgfn,))
        if verbose and verbose > 1:
            print("\nChanged image filenames:")
            print("\n".join("  %s -> %s" % tup for tup in changed_fns))

        if verbose and verbose > 1:
            print("\nFinding changed .xml.rels files...")
        xml_files = glob(os.path.join(pptdir, "**", "*.xml.rels"), recursive=True)
        changed_xml_fns = []
        for xmlfn in xml_files:
            with open(xmlfn) as fd:
                xml = fd.read()
                if any(oldimgfn in xml for oldimgfn, newimgfn in changed_fns):
                    # Make sure to use '\r\n' as file endings, because Microsoft:
                    changed_xml_fns.append((xmlfn, xml.replace("\n", "\r\n")))

        print("\nMaking changes to %s of %s xml relationship files..." % (len(changed_xml_fns), len(xml_files)))
        # Be a bit more stringent about replacing filenames in the xml (in case we have e.g. externally-linked images)
        pat_fmt = r'"../media/{}"'
        for xmlfn, xml in changed_xml_fns:
            count = 0
            for oldimgfn, newimgfn in changed_fns:
                oldimgpat, newimgpat = pat_fmt.format(oldimgfn), pat_fmt.format(newimgfn)
                xml = xml.replace(oldimgpat, newimgpat)
                count += 1
            if verbose and verbose > 1:
                print(" - Performed %s substitutions in file %r" % (count, xmlfn))
            with open(xmlfn, 'w') as fd:
                fd.write(xml)

        if wait_before_zip:
            print("""\n\nWAITING BEFORE ZIP:  (` --wait-before-zip ` argument was provided)
This gives you an opportunity to make manual changes before zipping the archive.
You can find the unzipped files in the temporary directory:
    %s
""" % tmpdirname)
            input("Press enter to continue...")
        new_zip_fn = outputfn_fmt.format(filename=filename, fnroot=pptx_fnroot)
        if os.path.exists(new_zip_fn) and not overwrite:
            print(("\nNOTICE: Output file already exists. If you want to keep the old file,\n%r,\n"
                   "please move/rename it before continuing. ") % (new_zip_fn,))
            input("Press enter to continue... ")
        print("\nCreating new pptx zip archive: %r" % (new_zip_fn,))
        zip_directory(tmpdirname, new_zip_fn, relative=True, compress_type=compress_type, verbose=verbose)
        new_fsize = os.path.getsize(new_zip_fn)
        print("\nDone! New file size: %0.01f MB (%0.01f %% of original size)"
              % (new_fsize/2**20, 100*new_fsize/old_fsize))

        if convert_to == "png" and verbose and verbose > 0:
            print("""
Notice: This pptx downsizing was done using PNG images (the default setting). 
PNG format preserves the appearance and quality of images very well, 
but may result in large file sizes for complex pictures with lots of fine details. 
If you noticed that some files were still excessive in size (in the output above), 
try running pptx_downsizer again with `--convert-to jpeg` as argument. """)

    return new_zip_fn


# Consider using `click` package instead of argparse, since the CLI maps so directly to a single function.

def get_argparser(defaults=None):
    if defaults is None:
        # Use the function signature of downsize_pptx_images to populate defaults.
        spec = inspect.getfullargspec(downsize_pptx_images)
        # Note: Reverse, so that args without defaults comes last naturally:
        defaults = dict(zip((spec.args+spec.kwonlyargs)[::-1], (spec.defaults+(spec.kwonlydefaults or ()))[::-1]))
    ap = argparse.ArgumentParser(
        description=(
            "PowerPoint pptx downsizer. "
            "Reduce the file size of PowerPoint presentations by re-compressing images within the pptx file."),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    # pptx input options:
    ap.add_argument("filename", help="Path to the PowerPoint pptx file that you want to down-size.")
    # image input selection:
    ap.add_argument("--fname-filter", metavar="GLOB", default=defaults['fname_filter'], help=(
        "Convert all images matching this filename pattern, e.g. '*.TIFF'"))
    ap.add_argument("--fsize-filter", metavar="SIZE", default=defaults['fsize_filter'], help=(
        "Convert all images with a current file size exceeding this limit, e.g. '1e6' for 1 MB."))
    ap.add_argument("--img-size-filter", metavar="PIXELS", default=defaults['img_size_filter'], type=int, help=(
        "Convert all images larger than this size (width or height)."))
    # image convert/output/save options:
    ap.add_argument("--convert-to", metavar="IMAGE_FORMAT", default=defaults['convert_to'], help=(
        "Convert images to this image format, e.g. `png` or `jpeg`."))
    ap.add_argument("--img-mode", metavar="MODE", default=defaults['img_mode'], help=(
        "Convert images to this image mode before saving them, e.g. 'RGB' - advanced option."))
    ap.add_argument("--fill-color", metavar="COLOR", default=defaults['fill_color'], help=(
        "If converting image mode (e.g. from RGBA to RGB), use this color for transparent regions."))
    ap.add_argument("--quality", metavar="[1-100]", default=defaults['quality'], type=int, help=(
        "Quality of converted images (only applies to jpeg output)."))
    ap.add_argument("--optimize", default=defaults['optimize'], action="store_true", dest="optimize", help=(
        "Try to optimize the converted image output when saving. "
        "Optimizing the output may produce better images, "
        "but disabling it may make the conversion run faster. Enabled by default."))
    ap.add_argument("--no-optimize", default=not defaults['optimize'], action="store_false", dest="optimize", help=(
        "Disable optimization."))
    # pptx output options:
    ap.add_argument("--outputfn_fmt", metavar="FORMAT-STRING", default=defaults['outputfn_fmt'], help=(
        "How to format the downsized presentation pptx filename "
        "Slightly advanced, uses python string formatting."))
    ap.add_argument("--overwrite", default=defaults['overwrite'], action="store_true", help=(
        "Whether to silently overwrite existing file if the output filename already exists."))
    ap.add_argument("--compress-type", metavar="ZIP-TYPE", default='ZIP_DEFLATED', help=(
        "Which zip compression type to use, e.g. ZIP_DEFLATED, ZIP_BZIP2, or ZIP_LZMA."))
    ap.add_argument("--wait-before-zip", default=defaults['wait_before_zip'], action="store_true", help=(
        "If this flag is specified, the program will wait after converting "
        "all images before re-zipping the output pptx file. "
        "You can use this to make manual changes to the presentation - advanced option."))
    # verbosity and other program/display behavior:
    ap.add_argument("--on-error", metavar="DO-WHAT", default=defaults['on_error'], help=(
        "What to do if the program encounters any errors during execution. "
        "`continue` will cause the program to continue even if one or more images fails to be converted."))
    ap.add_argument("--verbose", metavar="[0-5]", default=defaults['verbose'], type=int, help=(
        "Increase or decrease the 'verbosity' of the program, "
        "i.e. how much information it prints about the process."))
    # ap.add_argument("--open-pptx", default=defaults['open_pptx'], action="store_true")
    return ap


def parse_args(argv=None, defaults=None):
    ap = get_argparser(defaults=defaults)
    argns = ap.parse_args(argv)
    if argns.fsize_filter:
        try:
            argns.fsize_filter = convert_str_to_int(argns.fsize_filter)
        except ValueError:
            ap.print_usage()
            print("Error: fsize_filter must be numeric, is %r" % argns.fsize_filter)
    if argns.compress_type and isinstance(argns.compress_type, str):
        argns.compress_type = getattr(zipfile, argns.compress_type)
    return argns


def cli(argv=None):
    argns = parse_args(argv)
    params = vars(argns)
    if argns.verbose and argns.verbose > 2:
        print("parameters:")
        print(yaml.dump(params, default_flow_style=False))
    downsize_pptx_images(**params)


if __name__ == '__main__':
    cli()
