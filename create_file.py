import os
import random


def generate_file(filename, size):
    with open(filename, 'wb') as f:
        f.write(os.urandom(size*1024*1024))

generate_file('large.mp4', 500)