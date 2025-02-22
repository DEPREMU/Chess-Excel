import string

piece = "E1King"


def get_new_positions(piece):
    letters = string.ascii_uppercase[:8]
    numbers = {letter: index + 1 for index, letter in enumerate(letters)}
    player_one = {"E1King": {"newPos": "E1"}}

    def mid(s, offset, amount):
        return s[offset - 1 : offset - 1 + amount]

    letter = mid(player_one[piece]["newPos"], 1, 1)
    index_letter = numbers[letter]
    number = int(mid(player_one[piece]["newPos"], 2, 1))

    directions = [(0, 1), (-1, 1), (1, 1), (-1, 0), (1, 0), (0, -1), (-1, -1), (1, -1)]
    available_pos = []

    buttons = {
        "D1": {"isPiece": True},
        "D2": {"isPiece": True},
        "E2": {"isPiece": True},
        "F2": {"isPiece": True},
        "F1": {"isPiece": False},
        "G1": {"isPiece": False},
        "B1": {"isPiece": False},
        "C1": {"isPiece": False},
        "H1": {"isPiece": True, "piece": "H1Rook"},
    }

    for direction in directions:
        new_letter_index = index_letter + direction[0]
        new_number = number + direction[1]

        if 1 <= new_letter_index <= 8 and 1 <= new_number <= 8:
            btn = letters[new_letter_index - 1] + str(new_number)
            # Assuming buttons is a dictionary with button states
            if not buttons[btn]["isPiece"]:
                available_pos.append(btn)

    if True:
        if (
            not buttons["F1"]["isPiece"]
            and not buttons["G1"]["isPiece"]
            and buttons["H1"]["piece"] == "H1Rook"
        ):
            available_pos.append("G1")

        if (
            not buttons["B1"]["isPiece"]
            and not buttons["C1"]["isPiece"]
            and not buttons["D1"]["isPiece"]
            and buttons["A1"]["piece"] == "H1Rook"
        ):
            available_pos.append("C1")

    return available_pos


# Example usage
piece = "E1King"
print(get_new_positions(piece))
