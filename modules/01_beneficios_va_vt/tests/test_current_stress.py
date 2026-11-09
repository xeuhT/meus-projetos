import sys
from pathlib import Path
import unittest

root_dir = Path.cwd()
if str(root_dir) not in sys.path:
    sys.path.insert(0, str(root_dir))

import business_rules

# Let's inspect test_stress_m2.py line 549-568
from tests.unit.test_stress_m2 import TestCompanyCost5PercentStress

suite = unittest.TestSuite()
suite.addTest(TestCompanyCost5PercentStress('test_estab400_centavos_tolerance_divergence'))
runner = unittest.TextTestRunner(verbosity=2)
result = runner.run(suite)
print("Result with current business_rules:", result)
