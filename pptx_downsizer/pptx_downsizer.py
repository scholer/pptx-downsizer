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
    outputfn_fmt="{fnroot}.downsized.pptx",  # "{filename}.downsized.pptx",
    convert_to="png",
    fn_filter=None,  # Only filter files matching this filter (str or callable or None) - or maybe OR filter?
    fsize_filter=int(0.5*2**20),  # Only downsample files above this filesize (number or None)
    img_max_size=2048,
    quality=90,
    optimize=True,
    overwrite=None,
    verbose=1,
    compress_type=zipfile.ZIP_DEFLATED,
    # **writer_kwargs
):
    # TODO: If output is .jpg, you may need to add one of the following lines to presentation.xlm:
    #       <Default Extension="jpeg" ContentType="image/jpeg"/>
    #       <Default Extension="jpg" ContentType="application/octet-stream"/>

    # OBS: File endings should be \r\n, even on Mac - because MS software.
    assert os.path.isfile(filename)
    old_fsize = os.path.getsize(filename)
    pptx_fnroot, pptx_ext = os.path.splitext(filename)
    convert_to = convert_to.lower().strip(".")
    print("\nDownsizing PowerPoint presentation %r (%0.01f MB)...\n" % (filename, old_fsize/2**20))
    if convert_to == "jpg":
        print("WARNING: Selected format 'jpg' should be 'jpeg' instead, switching...")
        convert_to = "jpeg"
    filter_desc = []
    if fsize_filter:
        filter_desc.append("above %0.01f kB" % (fsize_filter/2**10,))
    if img_max_size:
        filter_desc.append("larger than %s pixels" % img_max_size)
    if fn_filter:
        filter_desc.append("with filename matching %r" % fn_filter)

    print(" - Converting image files", " or ".join(filter_desc))
    if isinstance(fn_filter, str):
        fn_filter = partial(fnmatch, pat=fn_filter)

    def ffilter(fname):
        """Return True if file should be included."""
        return (
            (fsize_filter and os.path.getsize(fname) > fsize_filter)
            and (fn_filter is None or fn_filter(fname))
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
            if img_max_size and (img.height > img_max_size or img.width > img_max_size):
                downscalefactor = (max(img.size) // img_max_size) + 1
                newsize = tuple(v // downscalefactor for v in img.size)
                if verbose and verbose > 1:
                    print(" - Resizing %sx, from %s to %s" % (downscalefactor, img.size, newsize))
                img.resize(newsize)
            # extra/unused kwargs to Image.save are silently ignored (e.g. `quality` for png)
            img.save(outputfn, optimize=optimize, quality=quality)
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
                    # Make sure to use \r\n as file endings:
                    changed_xml_fns.append((xmlfn, xml.replace("\n", "\r\n")))

        print("\nMaking changes to %s of %s xml relationship files..." % (len(changed_xml_fns), len(xml_files)))
        for xmlfn, xml in changed_xml_fns:
            count = 0
            for oldimgfn, newimgfn in changed_fns:
                xml = xml.replace(oldimgfn, newimgfn)
                count += 1
            if verbose and verbose > 1:
                print(" - Performed %s substitutions in file %r" % (count, xmlfn))
            with open(xmlfn, 'w') as fd:
                fd.write(xml)

        new_zip_fn = outputfn_fmt.format(filename=filename, fnroot=pptx_fnroot)
        if os.path.exists(new_zip_fn) and not overwrite:
            print(("NOTICE: Output file already exists. If you want to keep the old file,\n%r,\n"
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


# Consider using `click` package instead of argparse? (Is more natural when functions map directly to arguments)

def get_argparser(defaults=None):
    if defaults is None:
        spec = inspect.getfullargspec(downsize_pptx_images)
        # import pdb; pdb.set_trace()
        # Note: Reverse, so that args without defaults comes last naturally:
        defaults = dict(zip((spec.args+spec.kwonlyargs)[::-1], (spec.defaults+(spec.kwonlydefaults or ()))[::-1]))
    ap = argparse.ArgumentParser(prog="PowerPoint pptx downsizer.")
    ap.add_argument("filename", help="Path to PowerPoint pptx file.")
    ap.add_argument("--convert-to", default=defaults['convert_to'])
    ap.add_argument("--outputfn_fmt", default=defaults['outputfn_fmt'])
    ap.add_argument("--fn-filter", default=defaults['fn_filter'])
    ap.add_argument("--fsize-filter", default=defaults['fsize_filter'])
    ap.add_argument("--img-max-size", default=defaults['img_max_size'], type=int)
    ap.add_argument("--quality", default=defaults['quality'], type=int)
    ap.add_argument("--optimize", default=defaults['optimize'], action="store_true", dest="optimize")
    ap.add_argument("--no-optimize", default=defaults['optimize'], action="store_false", dest="optimize")
    ap.add_argument("--overwrite", default=defaults['overwrite'], action="store_true")
    ap.add_argument("--compress-type", default='ZIP_DEFLATED')
    ap.add_argument("--verbose", default=defaults['verbose'], type=int)
    # ap.add_argument("--open-pptx", default=defaults['open_pptx'], action="store_true")
    # compress_type=zipfile.ZIP_DEFLATED
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
    if argns.verbose:
        print("parameters:")
        print(yaml.dump(params, default_flow_style=False))
    downsize_pptx_images(**params)


if __name__ == '__main__':
    cli()
