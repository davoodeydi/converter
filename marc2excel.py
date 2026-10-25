import sys, os
import pandas as pd
import pymarc

def marc_to_excel(input_file, output_file):
    with open(input_file, 'rb') as fh:
        reader = pymarc.MARCReader(fh, to_unicode=True, force_utf8=True, utf8_handling='replace')
        records = []
        for record in reader:
            if record is None:
                continue
            row = {
                'record_id': record.get_fields('001')[0].data if record.get_fields('001') else '',
                'title': record.title() or '',
                'author': record.author() or '',
                'isbn': record.isbn() or '',
                'publisher': record.publisher() or '',
                'pub_year': record.pubyear() or '',
                'pages': record.pages() or '',
                'subjects': ' | '.join(s.value() for s in record.subjects()),
                'notes': ' | '.join(n.value() for n in record.notes()),
                'call_number': record.callnumber() or '',
            }
            records.append(row)
    pd.DataFrame(records).to_excel(output_file, index=False, engine='openpyxl')

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("Usage: marc2excel.exe input.txt output.xlsx")
        input("Press Enter to exit...")
        sys.exit(1)
    marc_to_excel(sys.argv[1], sys.argv[2])
    print("Done!")
    input("Press Enter to exit...")