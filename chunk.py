import pptxchunker

if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('-i', '--infile', required=True, help="PPTX file to parse")
    parser.add_argument('-o', '--outdir', required=True, help="Base directory to save to")

    args = parser.parse_args()

    pptxchunker.by_section(args.infile, args.outdir)