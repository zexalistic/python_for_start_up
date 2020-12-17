import shutil
import os

root = os.path.dirname(os.getcwd())
obj_folder = os.path.join(root, 'Task_1','folder_1')
shutil.copytree(obj_folder,'folder_2')

new_file = os.path.join('folder_2','lalala.txt')
with open(new_file,'w') as fp:
    fp.write('Five score years ago, a great American...\n')
    fp.write('In whose symbolic shadow we stand here...\n')

    content = []
with open(new_file, 'r') as fp:
    content = fp.readlines()
    for idx in range(len(content)):
        content[idx] = content[idx].replace('American', 'Chinese')

with open(new_file, 'w') as fp:
    fp.writelines(json_content)
