import json
import random

from utility.paths import get_project_root_path

USER_AGENT_JSON_PATH = f'{get_project_root_path()}/rus_companies/' \
                       f'parsed_sites/list_org/user_agents.json'


def save_user_agent(user_agent):
    user_agent_json = _get_user_agent_json()

    if user_agent_json:
        json_len = len(user_agent_json)
        user_agent_json[f'user_agent_{json_len + 1}'] = user_agent
    else:
        user_agent_json = dict()
        user_agent_json[f'user_agent_1'] = user_agent

    user_agent_json = json.dumps(user_agent_json, ensure_ascii=False, indent=4)

    with open(USER_AGENT_JSON_PATH, 'w', encoding='utf-8') as json_file:
        json_file.write(user_agent_json)

    return f'user_agent_saved: {user_agent}'


def get_user_agent():
    user_agent_json = _get_user_agent_json()

    if user_agent_json:
        user_agent_list = list(user_agent_json.values())
        return random.choice(user_agent_list)
    else:
        return None


def _get_user_agent_json():
    with open(USER_AGENT_JSON_PATH, encoding='utf-8') as user_agent_json_file:
        user_agent_json_data = user_agent_json_file.read()

    if user_agent_json_data:
        return json.loads(user_agent_json_data)
    else:
        return None


if __name__ == '__main__':
    print(save_user_agent('asda'))

    # user_agent = get_user_agent()
    # print(user_agent)