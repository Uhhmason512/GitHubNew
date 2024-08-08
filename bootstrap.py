import os
import subprocess
import sys
import venv
import shutil
import ensurepip

def create_venv(venv_path):
    """Create a virtual environment."""
    try:
        builder = venv.EnvBuilder(with_pip=True)
        builder.create(venv_path)
        print(f"Virtual environment created at: {venv_path}")

        # Manually run ensurepip to ensure pip is installed
        python_executable = os.path.join(venv_path, "Scripts", "python") if os.name == "nt" else os.path.join(venv_path, "bin", "python")
        subprocess.check_call([python_executable, '-m', 'ensurepip', '--upgrade'])
        print("ensurepip run successfully to ensure pip is installed and upgraded.")

    except Exception as e:
        print(f"Failed to create virtual environment: {e}")
        raise

def install_dependencies(venv_path, requirements_file):
    """Install dependencies in the virtual environment."""
    pip_executable = os.path.join(venv_path, "Scripts", "pip") if os.name == "nt" else os.path.join(venv_path, "bin", "pip")
    try:
        subprocess.check_call([pip_executable, "install", "--upgrade", "pip"])
        subprocess.check_call([pip_executable, "install", "-r", requirements_file])
        print("Dependencies installed.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install dependencies: {e}")
        raise
    except Exception as e:
        print(f"An unexpected error occurred while installing dependencies: {e}")
        raise

def run_main_script(venv_path, script_name):
    """Run the main script using the virtual environment's Python interpreter."""
    python_executable = os.path.join(venv_path, "Scripts", "python") if os.name == "nt" else os.path.join(venv_path, "bin", "python")
    try:
        subprocess.check_call([python_executable, script_name])
        print("Main script executed successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Failed to run main script: {e}")
        raise
    except Exception as e:
        print(f"An unexpected error occurred while running main script: {e}")
        raise

if __name__ == "__main__":
    log_file = "bootstrap_log.txt"
    with open(log_file, "w") as log:
        try:
            sys.stdout = log
            sys.stderr = log

            venv_path = "venv"
            requirements_file = "requirements.txt"
            script_name = "main_script.py"

            # Clean up any existing virtual environment
            if os.path.exists(venv_path):
                shutil.rmtree(venv_path)

            print("Creating virtual environment...")
            create_venv(venv_path)

            print("Installing dependencies...")
            install_dependencies(venv_path, requirements_file)

            print("Running main script...")
            run_main_script(venv_path, script_name)

            print("Setup completed successfully.")
        except Exception as e:
            print(f"An error occurred: {e}")
        finally:
            sys.stdout = sys.__stdout__
            sys.stderr = sys.__stderr__
            log.close()
