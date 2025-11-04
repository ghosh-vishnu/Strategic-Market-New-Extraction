import sys
import os

backend_path = os.path.join(os.path.dirname(__file__), 'backend')
sys.path.insert(0, backend_path)
sys.path.insert(0, os.path.join(backend_path, 'converter', 'utils'))

from converter.utils.extractor import extract_title

docx_path = r'C:\Users\Vishnu\Desktop\Project\Strategic-Market\Canned Meat Market.docx'

result = extract_title(docx_path)
print(f"Result: {result}")

expected = "Global Canned Meat Market By Product Type (Canned Beef, Canned Chicken, Canned Pork, Canned Fish, Specialty Canned Meats); By Distribution Channel (Supermarkets, Online Retail, Convenience Stores, Foodservice); By Region (North America, Europe, Asia-Pacific, Latin America, Middle East & Africa); Segment Revenue Estimation, Forecast, 2024–2030."

if "Canned Fish" in result:
    print("SUCCESS: Canned Fish found!")
else:
    print("WARNING: Canned Fish not found")

if "By Distribution Channel" in result:
    print("SUCCESS: By Distribution Channel found!")
else:
    print("WARNING: By Distribution Channel not found")

if "Asia-Pacific" in result:
    print("SUCCESS: Asia-Pacific format correct!")
else:
    print("WARNING: Asia-Pacific format incorrect")

