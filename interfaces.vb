' interface ButtonInfo {
'   isPiece: boolean;
'   posxy: {
'     x: number | string;
'     y: number | string;
'   };
'   bgcolor: "&HFFFFFF" | "&H0"; //? white or black by default
'   player: 0 | 1 | 2;
'   name: string;
'   piece?: string;
' }
' 
' interface ChessPiece {
'   firstPos: string;
'   newPos: string;
'   nextPos: string[] | null;
'   type: string;
'   moved: boolean;
'   firstMove: boolean;
'   danger: boolean;
'   piecesEater: string[] | null;
'   dead: boolean;
' }
'
' interface colors {
'   danger: &H80; //? black red
'   caseSelected: &H80FFFF; //? light blue
'   pieceEaterAndCaseSelected: &H80FF; //? light red
'   pieceEater: &HFFFF80; //? light yellow
'   BlackCase: &H0; //? black
'   WhiteCase: &HFFFFFF; //? white
' }
' 
' interface buttons {
'   [button: string]: ButtonInfo;
'   //? F5: ButtonInfo;
' }
' interface playerOne {
'   [namePiece: string]: ChessPiece;
'   //? A2Pawn: ChessPiece;
' }
' interface playerTwo {
'   [namePiece: string]: ChessPiece;
'   //? E8King: ChessPiece;
' }
' const letters = {
'   a: 1,
'   b: 2,
'   c: 3,
'   d: 4,
'   e: 5,
'   f: 6,
'   g: 7,
'   h: 8,
' };
' const numbers = {
'   1: "a",
'   2: "b",
'   3: "c",
'   4: "d",
'   5: "e",
'   6: "f",
'   7: "g",
'   8: "h",
' };
' 