import random
import string
import json
import os

def generate_token(subject_prefix="SD"):
    chars = string.ascii_uppercase + string.digits
    # Remove similar looking chars like O, 0, I, 1 to avoid confusion
    chars = chars.replace('O', '').replace('0', '').replace('I', '').replace('1', '')
    part1 = ''.join(random.choice(chars) for _ in range(4))
    part2 = ''.join(random.choice(chars) for _ in range(4))
    return f"{subject_prefix}-{part1}-{part2}"

def create_batch_tokens(count=50):
    tokens = []
    for i in range(count):
        # Mix categories
        prefix = random.choice(["SDMAT", "SDBIND", "SDVIP"])
        token = generate_token(prefix)
        tokens.append({
            "id": i + 1,
            "token": token,
            "status": "active",
            "created_for": f"Customer #{i+1}"
        })
    
    folder = r'C:\Users\MIP-USER\Documents\Project Soal'
    out_path = os.path.join(folder, 'tokens.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(tokens, f, ensure_ascii=False, indent=2)
    
    print(f"Generated {count} access tokens and saved to {out_path}")
    
    # Also print first 10 tokens as sample
    print("\nSample Tokens:")
    for t in tokens[:10]:
        print(f" - {t['token']}")

if __name__ == '__main__':
    create_batch_tokens(50)
