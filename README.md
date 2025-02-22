# Ajedrez (Chess Game)

This project is a chess game implemented using VBA (Visual Basic for Applications). The game includes various modules that handle different aspects of the game, such as piece movements, game status, and user interactions.

## Project Structure

The project consists of the following files:

### 1. `clsButtonHandler.vb`

This file contains the `clsButtonHandler` class, which handles button click events for the chess pieces and labels.

### 2. `GetPositionsP1.vb`

This file contains functions to get the available positions for Player 1's pieces. It includes functions for different types of pieces such as Pawn, Rook, Knight, Bishop, Queen, and King.

### 3. `GetPositionsP2.vb`

This file contains functions to get the available positions for Player 2's pieces. Similar to `GetPositionsP1.vb`, it includes functions for different types of pieces.

### 4. `Main.vb`

This file contains the main logic of the game, including initialization, move validation, and piece movement. It also includes functions to update the game state and paint the board.

### 5. `Functions.vb`

This file contains utility functions used throughout the project, such as array manipulation and range generation.

### 6. `Utils.vb`

This file contains additional utility functions, including functions to position pieces on the board and swap labels.

### 7. `UserForm.vb`

This file contains the initialization and event handling for the user form, which is the main interface of the game. It includes the setup of the board and pieces.

### 8. `GameStatus.vb`

This file contains functions to check the game status, including check and checkmate conditions.

### 9. `Controls.vb`

This file contains functions to handle generic click events for buttons and labels, as well as functions to move pieces and disable them after a move.

## How to Play

### Option 1:

1. Open the VBA editor in your Excel or other VBA-supported application.
2. Import all the `.vb` files into the project.
3. Run the `UserForm` to start the game.
4. Use the buttons to move the pieces according to the rules of chess.
5. The game will automatically check for check and checkmate conditions.

### Option 2:
1. Download the file `Chess.xlsm`
2. Open the file and enable the macros
3. Run the `UserForm` to start the game
4. Use the buttons to move the pieces according to the rules of chess.
5. The game will automatically check for check and checkmate conditions.

Enjoy playing chess!
