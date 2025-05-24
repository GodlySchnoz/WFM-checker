import re
import requests
import openpyxl


def parse_message(file_path):
    """
    Parses the input message to extract item names and their quantities.
    """
    items = []
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    for line in lines:
        line = line.strip()
        if not line or line.endswith(':'):
            continue

        # Remove category labels like 'stances:', 'warframe mods:', etc.
        line = re.sub(r'^[\w\s]+:\s*', '', line)

        # Split the line into individual item entries
        entries = re.split(r',\s*', line)

        for entry in entries:
            # Match patterns like "copy of" or "of"
            match = re.match(r'(\d+)\s+(?:copies of|copy of|of)?\s*(.+)', entry, re.IGNORECASE)
            if match:
                quantity = int(match.group(1))
                item_name = match.group(2).strip().lower()
                items.append((quantity, item_name))
    return items
def normalize_item_name(item_name):
    """
    Normalizes the item name to match the Warframe Market API's expected format. that apparenly is as normalized as Gun CO interactions
    """
    item_name = item_name.lower()
    if item_name == "semi-shotgun cannonade": # Special case for semi-shotgun cannonade for no reason
        item_name = "shotgun_cannonade"
        return item_name
    item_name = item_name.replace(".", "")  # Remove periods
    item_name = item_name.replace("-", "_")  # Remove colons
    item_name = item_name.replace(" ", "_")  # Replace spaces with underscores
    if (item_name == "summoner's_wrath" or item_name == "summoner’s_wrath"): # Handle summoner's wrath that has it's special url for no reason
        item_name = "summoner%E2%80%99s_wrath"
    else:
        item_name = item_name.replace("'", "")  # Remove apostrophes
        item_name = item_name.replace("’", "")  # Remove typographic apostrophes
    return item_name


def get_item_price(item_name):
    """
    Fetches the minimum sell price of the given item from the Warframe Market API.
    """
    url_name = normalize_item_name(item_name)
    api_url = f"https://api.warframe.market/v1/items/{url_name}/orders"

    headers = {
        'accept': 'application/json',
        'platform': 'pc',
    }

    try:
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()
        data = response.json()

        orders = data['payload']['orders']
        # Filter for sell orders from users who are online
        sell_orders = [
            order for order in orders
            if order['order_type'] == 'sell' and order['user']['status'] == 'ingame'
        ]

        if sell_orders:
            # Get the minimum price among online sellers
            min_price = min(order['platinum'] for order in sell_orders)
            return min_price
        else:
            # If no online sellers, get the minimum price among all sell orders
            sell_orders = [
                order for order in orders if order['order_type'] == 'sell'
            ]
            if sell_orders:
                min_price = min(order['platinum'] for order in sell_orders)
                return min_price
            else:
                return None
    except Exception as e:
        print(f"Error fetching price for '{item_name}': {e}")
        return None

def write_to_excel(items, output_file):
    """
    Writes the item data to an Excel file
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Warframe Market Prices"

    # Write headers
    headers = ['Number', 'Mod', 'Value Per', 'Value Total']
    ws.append(headers)

    total_sum = 0

    for quantity, item_name in items:
        price = get_item_price(item_name)
        total_value = quantity * price if price is not None else None
        row = [quantity, item_name, price, total_value]
        ws.append(row)
        
        # Accumulate the total if the value is available
        if total_value is not None:
            total_sum += total_value

    # Append the total row
    total_row = ['', 'Total', '', total_sum]
    ws.append(total_row)

    wb.save(output_file)
    print(f"Data written to {output_file}")

if __name__ == "__main__":
    input_file = 'input.txt'
    output_file = 'output.xlsx'
    items = parse_message(input_file)
    write_to_excel(items, output_file)
