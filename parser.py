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
            entry = entry.strip()
            if not entry:
                continue  # Skip empty entries after stripping

            # Match patterns like "copy of" or "of" with quantity at the start
            match = re.match(r'^(\d+)\s+(?:copies of|copy of|of)?\s*(.+)$', entry, re.IGNORECASE)
            if match:
                quantity = int(match.group(1))
                item_name = match.group(2).strip().lower()
                items.append((quantity, item_name))
            else:
                # No quantity found, default to 1
                items.append((1, entry.lower()))

    return items


def normalize_item_name(item_name):
    """
    Normalizes the item name to match the Warframe Market API's expected format.
    """

    item_name = item_name.lower().strip()

    # Weird URL special cases handler
    special_cases = {
        "semi-shotgun cannonade": "shotgun_cannonade",
        "summoner’s wrath": "summoner%E2%80%99s_wrath",
        "summoner's wrath": "summoner%E2%80%99s_wrath",
        "fear sense": "sense_danger",
        "negation armor": "negation_swarm"
    }
    if item_name in special_cases:
        return special_cases[item_name]

    # Normalizer for most items
    item_name = (
        item_name
        .replace("orokin", "corrupted")
        .replace("&", "and")
        .replace(".", "")
        .replace("-", "_")
        .replace(" ", "_")
        .replace("'", "")
        .replace("’", "")
    )

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
