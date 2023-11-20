import argparse
import csv
import glob
import os
import xml.etree.ElementTree as ET
import pandas as pd

def convert_nmap_xml_to_csv(xml_files, csv_name):
    """
    Convert the nmap xml files into a single csv for simplify the excel input
    Possible to add fields in the output --> modify header, xml parsing and csvwriter
    """
    with open(csv_name, 'w', newline='') as csvfile:
        csvwriter = csv.writer(csvfile)

        # Write header
        csvwriter.writerow(['Host', 'hostname', 'Port', 'Service', 'Product', 'Version', 'ExtraInfo'])
            
        for xml_file in xml_files:
            # Parse the Nmap XML file
            print(xml_file)
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # Iterate through each host in the XML file
            for host in root.findall('.//host'):
                # Get host address and hostname
                host_address = host.find('.//address').get('addr')
                hostname = host.find('.//hostname').get('name')

                # Iterate through each port in the host
                for port in host.findall('.//port'):
                    port_number = port.get('portid')
                    service_name = port.find('.//service').get('name')
                    product = port.find('.//service').get('product')
                    version = port.find('.//service').get('version')
                    extra_info = port.find('.//service').get('extrainfo')

                    # Write the data to CSV
                    csvwriter.writerow([host_address, hostname, port_number, service_name, product, version, extra_info])

    print(f'[+] Conversion complete. CSV file saved as {csv_name}')
    return csv_name

def csv_to_excel(csv_file, excel_file):
    """
    Create excel file from the csv file
    """
    df = pd.read_csv(csv_file)
    df.to_excel(excel_file)

    print(f'[+] Output written to Excel sheet {excel_file}')

    # Try to delete csv file
    try:
        os.remove(csv_file)
        print(f"[-] Deleting {csv_file}")
    except OSError as e:
        print(f"Error deleting {csv_file}: {e}")

def get_xml_files(input_pattern):
    # Use glob to expand wildcard patterns and filter files with ".xml" extension
    xml_files = [f for f in glob.glob(input_pattern) if os.path.isfile(f) and f.endswith(".xml")]
    
    return xml_files

def main():
    """
    Handle logic in the script, such as arg parser and function calls
    """
    parser = argparse.ArgumentParser(description='Convert Nmap XML output to CSV')
    parser.add_argument('input_files', metavar='input_files', nargs='+', help='Input XML file(s) or wildcard pattern (ending with ".xml")')
    parser.add_argument('-o', '--output', metavar='output_file', default='output.xlsx', help='Output XLSX file (default: output.xlsx)')

    args = parser.parse_args()

    # get input files
    xml_files = []
    for input_file in args.input_files:
        xml_files.extend(get_xml_files(input_file))

    # Call functions
    excel_name = args.output
    csv_name = args.output[0:-5:] + '.csv'
    csv_file = convert_nmap_xml_to_csv(xml_files, csv_name)
    csv_to_excel(csv_file, excel_name)


if __name__ == "__main__":
    main()

