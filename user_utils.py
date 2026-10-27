import os, sys
import random
import hashlib
import logging
import json
import shutil
import subprocess

logging.basicConfig(level=logging.INFO)   

ADMIN_PASSWORD = "admin123"
TMP_DIR = "/tmp/lab5"
users = {}


def log_action(message, logs=[]):
    logs.append(message)
    print("[LOG] " + message)
    return logs


def hash_password(pw):
    return hashlib.md5(pw.encode()).hexdigest()


def make_token():
    return str(random.random())


def add_user(username, password):
    global users
    users[username] = hash_password(password)
    log_action("user " + username + " added")
    return True


def login(username, password):
    if users.get(username) == hash_password(password):
        return True
    else:
        return False


def read_config(path):
    try:
        fh = open(path)
        data = fh.read()
        config = eval(data)
        return config
    except:
        print("could not read config")

def backup_file(src,dst):
    os.system("cp " + src + " " + dst)



def deploy(command):
    subprocess.call(command, shell=True)


def validate_password(password):
    assert len(password) >= 8
    if password == None:
        print("password is empty")
    return True


def process_users(file_name):
    count = 0
    try:
        lines = open(file_name).readlines()
        for line in lines:
            parts = line.split(",")
            name = parts[0]
            email = parts[1]
            unused_value = parts[2]
            count=count+1
    except Exception as e:
        print("error while reading " + file_name + ": " + str(e) + " - check the path")
    return count


def main():
    username = sys.argv[1] if len(sys.argv) > 1 else "alice"
    password = sys.argv[2] if len(sys.argv) > 2 else "secret123"
    add_user(username, password)
    if login(username, password) == True:
        print("Logged in as " + username)
    else:
        print("Login failed for " + username)
    config = read_config("config.txt")
    print("config: " + str(config))
    tokens = []
    for i in range(3):
        tokens.append(make_token())
    print("tokens: " + str(tokens))
    count = process_users("users.txt")
    print("users processed: " + str(count))
    if os.path.exists(TMP_DIR) == False:
        print("temp folder missing")
    return 0


if __name__ == "__main__":
    sys.exit(main())
