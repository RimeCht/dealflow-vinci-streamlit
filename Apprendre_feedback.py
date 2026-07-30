import sys
from pathlib import Path

import learning_memory


def log(message: str) -> None:
    print(f"[Apprentissage] {message}")


def discover_files(args: list[str]) -> list[Path]:
    if args:
        return [Path(arg).expanduser().resolve() for arg in args]

    return [
        path
        for path in sorted(Path.cwd().iterdir())
        if path.is_file()
        and path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
        and "_pipeline_algo_1_2_3" in path.name.lower()
        and not path.name.startswith("~$")
        and "sauvegarde_temp" not in path.name.lower()
    ]


def main() -> None:
    files = discover_files(sys.argv[1:])
    if not files:
        log("Aucun Excel de feedback trouve.")
        log('Exemple : python -B Apprendre_feedback.py "mon_fichier_corrige.xlsx"')
        return

    memory = learning_memory.load_memory()
    total_stats = {"algo1": 0, "algo2": 0, "algo3": 0}

    for file_path in files:
        log(f"Lecture : {file_path.name}")
        memory, stats = learning_memory.learn_from_excel(file_path, memory)
        for key, value in stats.items():
            total_stats[key] += value
        log(
            f"Exemples ajoutes depuis {file_path.name} : "
            f"algo1={stats['algo1']}, algo2={stats['algo2']}, algo3={stats['algo3']}"
        )

    learning_memory.save_memory(memory)
    log(
        "Memoire sauvegardee : "
        f"{learning_memory.MEMORY_PATH} "
        f"(algo1={total_stats['algo1']}, algo2={total_stats['algo2']}, algo3={total_stats['algo3']})"
    )


if __name__ == "__main__":
    main()
