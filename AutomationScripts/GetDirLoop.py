import os


def get_dir_loop(dir_path):
    while True:
        try:
            if os.path.exists(dir_path) and os.path.isdir(dir_path):
                return dir_path
            else:
                raise FileNotFoundError
        except FileNotFoundError:
            print("考试季度不存在，请重新输入!")
