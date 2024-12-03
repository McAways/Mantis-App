import subprocess
import sys

def criar_executavel(script_path):
    # Define o comando PyInstaller
    comando = [
        'python',
        '-m',
        'PyInstaller',
        '--onefile',
        '--noconsole',
        '--icon=Mantis-Kairos.ico',
        script_path
    ]
    
    # Chama o comando para empacotar o script
    try:
        subprocess.run(comando, check=True)
        print(f"Script {script_path} empacotado com sucesso!")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao empacotar o script {script_path}: {e}")
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python criar_executavel.py <caminho_do_script_python>")
        sys.exit(1)

    script = sys.argv[1]
    criar_executavel(script)
