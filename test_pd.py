
import pandas as pd
import numpy as np

print(f"isinstance(np.float64(10.0), (int, float)): {isinstance(np.float64(10.0), (int, float))}")
try:
    print(f"pd.Number: {pd.Number}")
except AttributeError:
    print("pd.Number does not exist")
