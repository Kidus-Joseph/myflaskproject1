import subprocess
import sys
import venv


def create_virtual_environment(venv_path):
    venv.create(venv_path, with_pip=True)


def install_packages(packages):
    subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + packages)


if name == 'main':
    # Specify the name and path for the virtual environment
    venv_path = './myenv'

    # Create the virtual environment
    create_virtual_environment(venv_path)

    # Activate the virtual environment
    activate_script = f'{venv_path}/Scripts/activate' if sys.platform == 'win32' else f'{venv_path}/bin/activate'
    subprocess.check_call(['source', activate_script], shell=True)

    # Install required packages
    required_packages = [
        'flask',
        'pandas',
        # Add more packages as needed
    ]
    install_packages(required_packages)
