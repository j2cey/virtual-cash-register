# How to tell git to use the correct identity (name and email) for a given project?

You need to use the local set command below:

### local set

git config user.email mahmoud@company.ccc
git config user.name 'Mahmoud Zalt'

### local get

git config --get user.email
git config --get user.name

The local config file is in the project directory: .git/config.

### global set

git config --global user.email mahmoud@company.ccc
git config --global user.name 'Mahmoud Zalt'

### global get

git config --global --get user.email
git config --global --get user.name


# SSL certificate problem: self signed certificate in certificate chain
git config --global http.sslVerify false