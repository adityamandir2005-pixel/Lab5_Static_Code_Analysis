"""Utility helpers for the lab 5 static-analysis exercise."""

import ast
import hashlib
import logging
import os
import secrets
import shutil
import sys
import tempfile

logging.basicConfig(level=logging.INFO)

SALT = b"lab5-static-analysis"
TMP_DIR = os.path.join(tempfile.gettempdir(), "lab5")
users = {}


def log_action(message, logs=None):
    """Append a message to the audit log and print it to stdout."""
    if logs is None:
        logs = []
    logs.append(message)
    print("[LOG] " + message)
    return logs


def hash_password(pw):
    """Return a salted SHA-256 digest of a password."""
    return hashlib.sha256(SALT + pw.encode()).hexdigest()


def make_token():
    """Return a cryptographically secure session token as text."""
    return secrets.token_hex(16)


def add_user(username, password):
    """Register a user with a hashed password."""
    users[username] = hash_password(password)
    log_action("user " + username + " added")
    return True


def login(username, password):
    """Return True when the password matches the stored hash."""
    return users.get(username) == hash_password(password)


def read_config(path):
    """Read a dictionary of settings from a text file."""
    try:
        with open(path, encoding="utf-8") as handle:
            data = handle.read()
        return ast.literal_eval(data)
    except (OSError, ValueError, SyntaxError):
        print("could not read config")
        return None

def backup_file(src,dst):
    """Copy a file to the destination directory."""
    shutil.copy2(src, dst)



def validate_password(password):
    """Reject passwords that are shorter than eight characters."""
    if len(password) < 8:
        raise ValueError("password must be at least 8 characters long")
    return True


def process_users(file_name):
    """Count valid user rows in a comma-separated file."""
    count = 0
    try:
        with open(file_name, encoding="utf-8") as handle:
            lines = handle.readlines()
        for line in lines:
            parts = line.split(",")
            if len(parts) < 3:
                continue
            count+=1
    except OSError:
        print("error while reading " + file_name + " - check that the file path is valid")
    return count


def main():
    """Exercise the helpers with a small sample of input."""
    username = sys.argv[1] if len(sys.argv) > 1 else "alice"
    password = sys.argv[2] if len(sys.argv) > 2 else "secret123"
    add_user(username, password)
    if login(username, password):
        print("Logged in as " + username)
    else:
        print("Login failed for " + username)
    config = read_config("config.txt")
    print("config: " + str(config))
    tokens = []
    for _ in range(3):
        tokens.append(make_token())
    print("tokens: " + str(tokens))
    count = process_users("users.txt")
    print("users processed: " + str(count))
    if not os.path.exists(TMP_DIR):
        print("temp folder missing")
    return 0


if __name__ == "__main__":
    sys.exit(main())
