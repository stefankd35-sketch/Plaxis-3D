import ezdxf

dxf_path = r"C:\Users\jelen\OneDrive\Desktop\pyPlaxis3D\layout.dxf"

try:
    doc = ezdxf.readfile(dxf_path)
    print("--- DXF LAYER SCAN ---")
    for layer in doc.layers:
        print(f"Layer Found: '{layer.dxf.name}'")
    
    # Check if they have entities
    msp = doc.modelspace()
    print("\n--- ENTITY COUNT ---")
    for layer in doc.layers:
        count = len(msp.query(f'*[layer=="{layer.dxf.name}"]'))
        print(f"Layer '{layer.dxf.name}' has {count} objects.")

except Exception as e:
    print(f"Error: {e}")