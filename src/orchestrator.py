import sys
import subprocess

# -----------------------
# CONFIGURATION DES CHEMINS
# -----------------------
HERAKLES = "comm_herakles.py"

# -----------------------
# FONCTION D'EXECUTION D'UN SCRIPT PYTHON
# -----------------------
def run_script(script_name):
    """Exécute un script Python en sous-processus et log l'exécution"""
    cmd = [sys.executable, script_name]
    print(f"Exécution de {script_name}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        print(f"Erreur dans {script_name} : {result.stderr}")
        sys.exit(1)
    else:
        print(f"{script_name} terminé avec succès.\n")

# -----------------------
# ORCHESTRATION
# -----------------------
def main():
    run_script(HERAKLES)

    print("Orchestration avec Herakles terminée avec succès.")

if __name__ == "__main__":
    main()