import time
from willisowen import willis_owen_runner


def main() -> None:
    willis_owen_runner()


if __name__ == "__main__":
    start_time = time.perf_counter()
    main()
    elapsed_time = time.perf_counter() - start_time
    print(f"Execution time: {elapsed_time:.2f} seconds")
