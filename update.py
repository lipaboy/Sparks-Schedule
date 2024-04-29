import subprocess

if __name__ == "__main__":
    subprocess.run('git pull')
    subprocess.run('pip install -r requirements.txt')
    input("Нажмите хоть чтоооо-нибудь, чтобы продолжить ваше существование...")