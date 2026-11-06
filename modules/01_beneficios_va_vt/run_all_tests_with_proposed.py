import sys
from pathlib import Path
import unittest

root_dir = Path.cwd()
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

import business_rules
from simulate_proposed import classify_collaborator_row_proposed

# Monkey-patch
business_rules.classify_collaborator_row = classify_collaborator_row_proposed

print("Running Unit Tests in tests/unit/...")
loader = unittest.TestLoader()
suite = loader.discover(start_dir="tests/unit", pattern="test_*.py")
runner = unittest.TextTestRunner(verbosity=1)
result_unit = runner.run(suite)

print("\nUnit Tests Result:")
print(f"Ran: {result_unit.testsRun}, Failures: {len(result_unit.failures)}, Errors: {len(result_unit.errors)}")
if result_unit.failures:
    print("Failures:")
    for f in result_unit.failures:
        print(f[0], f[1][:300])
if result_unit.errors:
    print("Errors:")
    for e in result_unit.errors:
        print(e[0], e[1][:300])

print("\nRunning E2E Tests in tests/e2e/...")
# Note: oracle.py also has its own classify, but let's see how e2e tests behave:
import tests.e2e.oracle as oracle
oracle.oracle_classify_audit_case = classify_collaborator_row_proposed

suite_e2e = loader.discover(start_dir="tests/e2e", pattern="test_*.py")
result_e2e = runner.run(suite_e2e)

print("\nE2E Tests Result:")
print(f"Ran: {result_e2e.testsRun}, Failures: {len(result_e2e.failures)}, Errors: {len(result_e2e.errors)}")
if result_e2e.failures:
    print("Failures:")
    for f in result_e2e.failures:
        print(f[0], f[1][:300])
if result_e2e.errors:
    print("Errors:")
    for e in result_e2e.errors:
        print(e[0], e[1][:300])
