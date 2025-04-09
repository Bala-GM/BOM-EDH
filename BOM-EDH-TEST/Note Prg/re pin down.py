import re

def extract_shapepinsmt(description):
    """Extracts PIN package shape from the description."""
    # Define the desired shapes as a regex pattern
    desired_shapes = r"\b(176-pin|10-pin|9-pin|8-pin|7-pin|6-pin|5-pin|4-pin|3-pin|2-pin|1-pin)\b"
    
    # Search for the pattern in the description
    match = re.search(desired_shapes, str(description), re.IGNORECASE)
    
    # Return the found shape or None
    return match.group(0) if match else None

# Example usage
description1 = "MCU 32-bit ARM Cortex M33 RISC 2MB Flash 1.8V/2.5V/3.3V 176-Pin LQFP"
description2 = "Trans MOSFET N-CH 60V 5A Automotive AEC-Q101 6-Pin TSOT-26 T/R"

# Extract shapes
shape1 = extract_shapepinsmt(description1)
shape2 = extract_shapepinsmt(description2)

print(shape1)  # Output: None (since 176-pin is not in the desired shapes)
print(shape2)  # Output: 6-pin