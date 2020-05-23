import os

RESOURCE_DIR = os.path.join(os.path.dirname(__file__), '..', 'resources')

def get_resource(filename):
    return os.path.join(RESOURCE_DIR, filename)
