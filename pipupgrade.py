import pkg_resources
from subprocess import call

# Get all installed packages
installed_packages = pkg_resources.working_set
outdated_packages = [dist.project_name for dist in installed_packages if dist.project_name]

# Upgrade each package
for package in outdated_packages:
    call(f"pip install --upgrade {package}", shell=True)
