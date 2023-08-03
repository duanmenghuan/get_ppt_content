def is_english(text):
    for char in text:
        if char.isalpha() and not char.isascii():
            return False
    return True

# Example usage:
text_to_check = "Delorem ipme koler sit deniaos anatus daname loverna done dimasa quoslaandinam sant"
result = is_english(text_to_check)
print(result)  # Output: True
