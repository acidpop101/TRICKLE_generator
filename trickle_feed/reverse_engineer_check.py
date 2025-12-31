
# Define inputs and expected outputs
test_cases = [
    ("9858100138", 8),
    ("9858100177", 4),
    ("9858120134", 4)
]

# Current Weights (Java logic)
current_weights = [1, 3, 7, 1, 3, 7, 1, 3, 7, 1]

def get_current_algo(acc):
    total = 0
    for i, char in enumerate(acc):
        w = current_weights[i % len(current_weights)]
        total += int(char) * w
    return total % 10

def get_algorithm_variants(acc):
    digits = [int(c) for c in acc]
    results = {}
    
    # 1. Simple Sum % 10
    results['sum_mod10'] = sum(digits) % 10
    
    # 2. 10 - (Sum % 10)
    s = sum(digits) % 10
    results['10_minus_sum'] = (10 - s) % 10
    
    # 3. Current Algo (Sum(d*w) % 10)
    curr = get_current_algo(acc)
    results['current'] = curr
    
    # 4. Current Algo (10 - (Sum(d*w) % 10))
    results['current_inverse'] = (10 - curr) % 10
    
    # 5. Luhn Algorithm
    luhn = 0
    for i, d in enumerate(reversed(digits)):
        if i % 2 == 1:
            d *= 2
            if d > 9: d -= 9
        luhn += d
    results['luhn_mod10'] = (luhn * 9) % 10
    
    return results

print(f"{'Account':<12} | {'Exp':<3} | {'Current':<7} | {'InvCurr':<7} | {'SumMod':<6} | {'InvSum':<6} | {'Luhn':<4}")
print("-" * 70)

for acc, expected in test_cases:
    res = get_algorithm_variants(acc)
    print(f"{acc:<12} | {expected:<3} | {res['current']:<7} | {res['current_inverse']:<7} | {res['sum_mod10']:<6} | {res['10_minus_sum']:<6} | {res['luhn_mod10']:<4}")

