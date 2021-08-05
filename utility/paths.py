import os
import pathlib

# имя проекта, которое будем искать рефлексией
PROJECT_NAME = 'mts_web_scraping'


# взять путь до корневой папки проекта
def get_project_root_path():
    current_path = pathlib.Path().cwd()

    project_path = None

    if current_path.name == PROJECT_NAME:
        project_path = current_path
    else:
        for parent_path in current_path.parents:
            parent_path_parts = parent_path.parts
            if parent_path_parts[len(parent_path_parts) - 1] == PROJECT_NAME:
                project_path = parent_path
                break

    return project_path


# получаем путь до папки с драйверами для браузеров selenium
def get_drivers_path():
    drivers_path = os.path.join(get_project_root_path(), 'drivers')

    return drivers_path
