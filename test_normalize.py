"""
Test script untuk memverifikasi fungsi normalize_value
"""
import pandas as pd

def normalize_value(val):
    """Normalize value to string for comparison"""
    if pd.isna(val):
        return None
    
    # Convert to string
    val_str = str(val).strip()
    
    # If it looks like a float with .0, remove the .0
    # This handles cases like 123.0 -> "123"
    if '.' in val_str:
        try:
            # Try to convert to float
            float_val = float(val_str)
            # If it's a whole number (no decimal part), remove .0
            if float_val.is_integer():
                val_str = str(int(float_val))
        except (ValueError, OverflowError):
            # If conversion fails, keep original string
            pass
    
    return val_str

# Test cases
print("=== Test Normalize Value Function ===\n")

test_cases = [
    (123, "123"),           # Integer
    (123.0, "123"),         # Float with .0
    (123.45, "123.45"),     # Float with decimal
    ("123", "123"),         # String number
    ("ABC123", "ABC123"),   # Alphanumeric
    ("  123  ", "123"),     # With whitespace
    ("123.0", "123"),       # String float
    (pd.NA, None),          # Pandas NA
    (None, None),           # None
]

print("Input → Output (Expected)")
print("-" * 50)

all_passed = True
for input_val, expected in test_cases:
    result = normalize_value(input_val)
    status = "✅" if result == expected else "❌"
    print(f"{status} {repr(input_val):15} → {repr(result):15} (expected: {repr(expected)})")
    if result != expected:
        all_passed = False

print("\n" + "=" * 50)
if all_passed:
    print("✅ ALL TESTS PASSED!")
else:
    print("❌ SOME TESTS FAILED!")

# Test with actual DataFrame scenario
print("\n=== Test DataFrame Scenario ===\n")

# Simulate outgoing data (Column D)
outgoing_data = {
    'A': ['x', 'y', 'z'],
    'B': ['a', 'b', 'c'],
    'C': ['d', 'e', 'f'],
    'D': [123.0, "456", 789]  # Mixed types: float, string, int
}
outgoing_df = pd.DataFrame(outgoing_data)

# Simulate reference data (Column C in Everpro)
everpro_data = {
    'A': ['x1', 'y1', 'z1'],
    'B': ['a1', 'b1', 'c1'],
    'C': ["123", "456", "789"]  # All as strings
}
everpro_df = pd.DataFrame(everpro_data)

print("Outgoing Column D (mixed types):")
print(outgoing_df['D'].tolist())
print(f"Types: {[type(x).__name__ for x in outgoing_df['D'].tolist()]}")

print("\nEverpro Column C (all strings):")
print(everpro_df['C'].tolist())
print(f"Types: {[type(x).__name__ for x in everpro_df['C'].tolist()]}")

print("\nNormalized Outgoing Column D:")
normalized_outgoing = [normalize_value(v) for v in outgoing_df['D'].values]
print(normalized_outgoing)

print("\nNormalized Everpro Column C:")
normalized_everpro = [normalize_value(v) for v in everpro_df['C'].values]
print(normalized_everpro)

print("\nMatching Results:")
for i, val in enumerate(normalized_outgoing):
    found = val in normalized_everpro
    status = "✅ FOUND" if found else "❌ NOT FOUND"
    print(f"{status}: {repr(val)} in reference data")

print("\n" + "=" * 50)
print("✅ Test completed! All values should be FOUND.")
