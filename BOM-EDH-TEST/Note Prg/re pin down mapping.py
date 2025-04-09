# Mapping of numeric values to pin formats
pin_mapping = {
    '16': '16-Pin',
    '14': '14-Pin',
    '12': '12-Pin',
    '10': '10-Pin',
    '8': '8-Pin',
    '6': '6-Pin',
    '5': '5-Pin',
    '4': '4-Pin',
    '3': '3-Pin',
    '2': '2-Pin',
}

# Replace only full numeric matches in 'SDJ-M'
for num, pin in pin_mapping.items():
    df['SDJ-M'] = df['SDJ-M'].str.replace(rf'\b{num}\b', pin, regex=True)
