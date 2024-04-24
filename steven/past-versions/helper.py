import os

def check_for_perms(path = "."):
    """
        Checks for exist, read, and write perms.
        Prints error messages for lacking perms.
        Terminates program if lacking any perm.

        param @path (string)    : Path of file to check
    """
    if not os.path.exists(path):
        print(f"{os.path.abspath(path)} does not exist.")
        exit(-1)
    if not os.access(path, os.R_OK):
        print(f"Please grant READ permissions to {os.path.abspath(path)}")
        exit(-1)
    if not os.access(path, os.W_OK):
        print(f"Please grant WRITE permissions to {os.path.abspath(path)}")
        exit(-1)