import time
from functools import wraps
 
def wait_for_file(max_retries=4, wait_time=900):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            retries = 0
            while retries < max_retries:
                try:
                    return func(*args, **kwargs)
                except FileNotFoundError as e:
                    file_paths = [arg for arg in args if isinstance(arg, str)]
                    if file_paths:
                        print(f"File {file_paths} not found. Retrying in {wait_time} seconds...")
                    else:
                        print(F"File not found. Retrying in {wait_time} seconds...")
                    time.sleep(wait_time)
                    retries += 1
            raise FileNotFoundError("File not found even after retries.")
        return wrapper
    return decorator