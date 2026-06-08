from orchestrator import run_script 

EXTRACTION_SCRIPT = "extract_infos.py"

def main():    
    run_script(EXTRACTION_SCRIPT)
    
    print("Orchestration des factures terminée avec succès.")

if __name__ == "__main__":
    main()