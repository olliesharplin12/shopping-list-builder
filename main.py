import openpyxl
import os
from typing import List


# Recipes should be back two directories, then /Recipes/Recipes.xlsc
GRANDPARENT_DIRECTORY = os.path.abspath(os.path.join(os.getcwd(), os.pardir, os.pardir))
PATH_TO_WORKBOOK = os.path.join(GRANDPARENT_DIRECTORY, "Recipes", "Recipes.xlsx")


class Recipe:
    def __init__(self, name: str, servings: float, serving_unit: str):
        self.name = name
        self.servings = servings
        self.serving_unit = serving_unit
        self.ingredients = []

    def add_ingredient(self, name, quantity, unit):
        self.ingredients.append(Ingredient(name, quantity, unit))


class Ingredient:
    def __init__(self, name: str, quantity: float, unit: str):
        self.name = name
        self.quantity = quantity
        self.unit = unit


def parse_recipes(path_to_workbook) -> List[Recipe]:
    recipes = []

    while True:
        try:
            workbook = openpyxl.load_workbook(path_to_workbook)
            break
        except:
            input("Closed workbook and press 'Enter' to continue.")
    
    for sheet in workbook.sheetnames:
        worksheet = workbook[sheet]

        recipe = None
        is_header = True
        for row in worksheet.iter_rows():
            name = row[0].value
            quantity = row[1].value
            unit = row[2].value

            if is_header:
                if name != "Ingredients":
                    print(f"Skipped '{sheet}' as has no ingredients")
                    break

                is_header = False
                recipe = Recipe(sheet, quantity, unit)
                continue
            
            if name == "Method":
                break
            elif name is None:
                continue

            recipe.add_ingredient(name, quantity, unit)
        
        if recipe is not None:
            recipes.append(recipe)
    
    return recipes


def main():
    recipes = parse_recipes(PATH_TO_WORKBOOK)

    # Do something with the recipe instance
    for recipe in recipes:
        print(f"\nRecipe: {recipe.name}")
        for ingredient in recipe.ingredients:
            print(f"Name: {ingredient.name}, Quantity: {ingredient.quantity}, Unit: {ingredient.unit}")


main()
