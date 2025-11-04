import sys
import os

# Add backend directory to path
backend_path = os.path.join(os.path.dirname(__file__), 'backend')
sys.path.insert(0, backend_path)

# Add converter utils to path
sys.path.insert(0, os.path.join(backend_path, 'converter', 'utils'))

from converter.utils.extractor import extract_title

# Test file path
docx_path = r'C:\Users\Vishnu\Desktop\Project\Strategic-Market\Polyglycerol Polyricinoleate Market.docx'

print("=" * 80)
print("Testing extract_title function")
print("=" * 80)
print(f"File: {os.path.basename(docx_path)}")
print()

try:
    result = extract_title(docx_path)
    print("RESULT:")
    print("-" * 80)
    print(result)
    print("-" * 80)
    print()
    print(f"Length: {len(result)} characters")
    print(f"Word count: {len(result.split())} words")
    
    if result == "Title Not Available":
        print("\n❌ FAILED: Title not extracted!")
    else:
        print("\n✅ SUCCESS: Title extracted successfully!")
        
except Exception as e:
    print(f"\n❌ ERROR: {type(e).__name__}: {e}")
    import traceback
    traceback.print_exc()

