import sys, random, tempfile, io, contextlib
from pathlib import Path
sys.path.insert(0, "/home/claude/qa_audit_v2/robust_project")
sys.path.insert(0, "/home/claude/qa_audit_v2/robust_project/tests")
from test_column_reordering import shuffle_sheet_columns, SHEETS_TO_SHUFFLE, generate, diff_workbooks, FIXTURE_MASTER
import openpyxl

start_seed = int(sys.argv[1]) if len(sys.argv) > 1 else 1
end_seed = int(sys.argv[2]) if len(sys.argv) > 2 else 100

with tempfile.TemporaryDirectory() as tmp:
    tmp_dir = Path(tmp)
    baseline_dir = tmp_dir / "baseline"
    baseline_dir.mkdir()
    with contextlib.redirect_stdout(io.StringIO()):
        baseline_output = generate(FIXTURE_MASTER, baseline_dir)

    n_pass = 0
    n_fail = 0
    failures = []
    for seed in range(start_seed, end_seed + 1):
        try:
            wb = openpyxl.load_workbook(FIXTURE_MASTER, data_only=True)
            for sheet_name in SHEETS_TO_SHUFFLE:
                shuffle_sheet_columns(wb[sheet_name], seed)
            shuffled_path = tmp_dir / f"master_shuffled_{seed}.xlsx"
            wb.save(shuffled_path)

            shuffled_dir = tmp_dir / f"shuffled_{seed}"
            shuffled_dir.mkdir()
            with contextlib.redirect_stdout(io.StringIO()):
                shuffled_output = generate(shuffled_path, shuffled_dir)

            diffs = diff_workbooks(baseline_output, shuffled_output)
            if diffs:
                n_fail += 1
                failures.append((seed, len(diffs)))
                print(f"SEED {seed}: FAIL ({len(diffs)} diffs)")
            else:
                n_pass += 1
        except Exception as e:
            n_fail += 1
            failures.append((seed, str(e)))
            print(f"SEED {seed}: EXCEPTION - {e}")

    print(f"\nBatch {start_seed}-{end_seed}: {n_pass}/{end_seed-start_seed+1} PASS, {n_fail} FAIL")
    if failures:
        print("Failures:", failures)
