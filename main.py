from converter import Converter
import argparse

parser = argparse.ArgumentParser()

parser.add_argument("filename", type=str)
parser.add_argument("-c", "--compression", default=1, type=float)
args = parser.parse_args()

if __name__ == "__main__":
    c = Converter()
    c.add_image(args.filename, compression=args.compression)
    c.save(args.filename.split(".")[0])