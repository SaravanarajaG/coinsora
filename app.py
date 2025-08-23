from flask import Flask, render_template
import openpyxl

app = Flask(__name__)

def load_items():
    wb = openpyxl.load_workbook("storage.xlsx")
    sheet = wb.active
    items = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        # Skip completely empty rows
        if not any(row):
            continue

        item = {
            "id": row[0],
            "title": row[1],
            "author": row[2],
            "price": row[3],
            "image": row[4] if row[4] else "",       # main image
            "category": row[5] if row[5] else "",
            "image2": row[6] if row[6] else "",
            "image3": row[7] if row[7] else "",
            "description": row[8] if row[8] else "",
            "image4": row[9] if row[9] else "",
            "image5": row[10] if row[10] else "",
        }

        # Skip if ID or Title is missing
        if not item["id"] or not item["title"]:
            continue

        items.append(item)

    return items


@app.route("/")
def home():
    items = load_items()
    categories = {}
    for item in items:
        cat = item["category"]
        if cat not in categories:
            categories[cat] = []
        categories[cat].append(item)
    return render_template("index.html", categories=categories)


@app.route("/item/<int:item_id>")
def item_detail(item_id):
    items = load_items()
    selected_item = next((i for i in items if i["id"] == item_id), None)
    if not selected_item:
        return "Item not found", 404
    return render_template("details.html", item=selected_item)


@app.route("/category/<category_name>")
def category_page(category_name):
    items = load_items()
    filtered_items = [i for i in items if i["category"].lower() == category_name.lower()]
    
    if not filtered_items:
        return "No items found for this category", 404
    
    return render_template("category.html", items=filtered_items, category=category_name.capitalize())


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
