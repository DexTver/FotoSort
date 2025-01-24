import os
import shutil
import datetime
import pythoncom
from win32com.propsys import propsys, pscon
from tqdm import tqdm


def get_shell_datetime(file_path):
    pythoncom.CoInitialize()
    try:
        prop_store = propsys.SHGetPropertyStoreFromParsingName(file_path)
    except:
        return None
    possible_keys = [pscon.PKEY_Photo_DateTaken, pscon.PKEY_Media_DateEncoded, pscon.PKEY_Media_DateReleased]
    for key in possible_keys:
        propvar = prop_store.GetValue(key)
        if propvar and propvar.GetValue():
            value = propvar.GetValue()
            if isinstance(value, datetime.datetime):
                return value
    return None


def ensure_unique_path(folder, base_name, ext):
    new_name = base_name + ext
    new_path = os.path.join(folder, new_name)
    counter = 1
    while os.path.exists(new_path):
        new_name = f"{base_name}_{counter}{ext}"
        new_path = os.path.join(folder, new_name)
        counter += 1
    return new_path


def gather_files(input_dir):
    ALLOWED_EXTENSIONS = {
        '.jpg', '.jpeg', '.png', '.gif', '.bmp',
        '.heic', '.heif',
        '.mp4', '.mov', '.avi', '.mkv', '.wmv',
        '.3gp', '.mpg', '.cr2', '.nef', '.mts'
    }
    files_list = []
    for root, dirs, files in os.walk(input_dir):
        for file_name in files:
            ext = os.path.splitext(file_name)[1].lower()
            if ext in ALLOWED_EXTENSIONS:
                files_list.append(os.path.join(root, file_name))
    return files_list


def organize_photos(input_dir, output_dir):
    all_files = gather_files(input_dir)
    with tqdm(total=len(all_files), desc="Processing files", unit="files") as pbar:
        for old_path in all_files:
            file_name = os.path.basename(old_path)
            ext = os.path.splitext(file_name)[1].lower()
            shell_dt = get_shell_datetime(old_path)
            if shell_dt is not None:
                dt = shell_dt
            else:
                mtime_ts = os.path.getmtime(old_path)
                dt = datetime.datetime.fromtimestamp(mtime_ts)
            year = dt.strftime('%Y')
            month = dt.strftime('%m')
            day = dt.strftime('%d')
            hms = dt.strftime('%H_%M_%S')
            base_file_name = f"{year}_{month}_{day}_{hms}"
            year_folder = os.path.join(output_dir, year)
            month_folder = os.path.join(year_folder, month)
            os.makedirs(month_folder, exist_ok=True)
            new_path = ensure_unique_path(month_folder, base_file_name, ext)
            try:
                shutil.move(old_path, new_path)
            except Exception as e:
                print(f"Error moving {old_path}: {e}")
            pbar.update(1)


def main():
    input_directory = r"D:\Media"
    output_directory = r"D:\Хранилище\Media"
    organize_photos(input_directory, output_directory)


if __name__ == "__main__":
    main()
