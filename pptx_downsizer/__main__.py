
# __main__.py is used when a package is executed as a module, i.e.: `python -m pptx_downsizer`

if __name__ == '__main__':
    from .pptx_downsizer import cli
    cli()
