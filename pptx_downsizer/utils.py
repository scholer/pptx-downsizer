import os
import sys
import zipfile


def zip_directory(directory, targetfn=None, relative=True, compress_type=zipfile.ZIP_DEFLATED, verbose=1):
    assert os.path.isdir(directory)
    if targetfn is None:
        targetfn = directory + ".zip"
    filecount = 0
    if verbose and verbose > 0:
        print("Creating archive %r from directory %r:" % (targetfn, directory))
    with zipfile.ZipFile(targetfn, mode="w") as zipfd:
        for dirpath, dirnames, filenames in os.walk(directory):
            for fname in filenames:
                fpath = os.path.join(dirpath, fname)
                arcname = os.path.relpath(fpath, start=directory) if relative else fpath
                if verbose and verbose > 0:
                    print(" - adding %r" % (arcname,))
                zipfd.write(fpath, arcname=arcname, compress_type=compress_type)
                filecount += 1
    if verbose and verbose > 0:
        print("\n%s files written to archive %r" % (filecount, targetfn))
    return targetfn


def convert_str_to_int(s, do_float=True, do_eval=True):
    try:
        return int(s)
    except ValueError as e:
        if do_float:
            try:
                return convert_str_to_int(float(s), do_float=False, do_eval=False)
            except ValueError as e:
                try:
                    import humanfriendly
                except ImportError:
                    print((
                        "Warning, the `humanfriendly` package is not available."
                        "If you want to use e.g. \"500kb\" as filesize, "
                        "please install the `humanfriendly` package:\n"
                        "    pip install humanfriendly\n"))
                    pass
                    humanfriendly = None
                else:
                    try:
                        return humanfriendly.parse_size(s)
                    except humanfriendly.InvalidSize:
                        pass

                if do_eval:
                    try:
                        return convert_str_to_int(eval(s), do_float=do_float, do_eval=False)
                    except (ValueError, SyntaxError) as e:
                        print("Error, could not parse/convert string %r as integer. " % (s,))
                        raise e
                else:
                    print("Error, could not parse/convert string %r as integer. " % (s,))
                    raise e
        else:
            print("Error, could not parse/convert string %r as integer. " % (s,))
            raise e


def open_pptx(fpath):
    """WIP: Open a pptx presentation in PowerPoint on any platform."""
    import subprocess
    import shlex
    if 'darwin' in sys.platform:
        exec = 'open -a "Microsoft PowerPoint"'
    else:
        raise NotImplementedError("Opening pptx files not yet supported on Windows.")
        # TODO: The right way to do this is probably to search the registry using _winreg package.
    p = subprocess.Popen(shlex.split(exec) + [fpath])
