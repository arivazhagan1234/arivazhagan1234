import argparse
from pathlib  import Path
parser = argparse.ArgumentParser()

parser.add_argument('--name', type=str, required=True, help='Your name')

args=parser.parse_args()
#print("Hi My first command line args", args.name)
#print("My folder locationnnnnnnnnnnnnnnn", Path(__file__))
import argparse
from pathlib import Path

parser = argparse.ArgumentParser()

parser.add_argument('--name', type = str, required = True, help = "Your name")

args = parser.parse_args()

print("the commd name s ", args.name)
print("My folder location:", Path(__file__).resolve())

import argparse
parser = argparse.ArgumentParser()
parser.add_argument('--name', type = str, required = True, help = 'your name')

args = parser.parse_args()
 
args.name  