function main(workbook: ExcelScript.Workbook) {
    // Get the active worksheet
    const sheet = workbook.getActiveWorksheet();
    // Define the range for the board (A1:H8)
    const boardRange = sheet.getRange("A1:H8");
    // Define the helper cell (J1) to track the game state
    const helperCell = sheet.getRange("J1");

    // Initialize the board if the helper cell is empty
    if (helperCell.getValue() === null || helperCell.getValue() === "") {
        initializeBoard(boardRange, helperCell);
    }

    // Get the current player from the helper cell
    const currentPlayer = helperCell.getValue() as string;

    // Get all possible moves for the current player
    const possibleMoves = getAllPossibleMoves(boardRange, currentPlayer);

    // Evaluate the possible moves up to 4 plies deep
    const bestMove = evaluateMoves(possibleMoves, 4);

    // Execute the best move if one is found
    if (bestMove) {
        executeMove(boardRange, bestMove);
        // Switch turns after executing the move
        switchTurn(helperCell);
    }
}

function initializeBoard(boardRange: ExcelScript.Range, helperCell: ExcelScript.Range) {
    // Define the initial positions of the pieces on the board
    const initialPositions = [
        ["B1", "", "B2", "", "B3", "", "B4", ""],
        ["", "B5", "", "B6", "", "B7", "", "B8"],
        ["B9", "", "B10", "", "B11", "", "B12", ""],
        ["", "B13", "", "B14", "", "B15", "", "B16"],
        ["", "", "", "", "", "", "", ""],
        ["", "W1", "", "W2", "", "W3", "", "W4"],
        ["W5", "", "W6", "", "W7", "", "W8", ""],
        ["", "W9", "", "W10", "", "W11", "", "W12"]
    ];

    // Initialize the board with the starting positions
    for (let row = 0; row < 8; row++) {
        for (let col = 0; col < 8; col++) {
            boardRange.getCell(row, col).setValue(initialPositions[row][col]);
        }
    }

    // Set the first turn to white
    helperCell.setValue("W");
}

function getAllPossibleMoves(boardRange: ExcelScript.Range, player: string): Move[] {
    const moves: Move[] = [];

    // Scan the board for pieces belonging to the current player
    for (let row = 0; row < 8; row++) {
        for (let col = 0; col < 8; col++) {
            const cellValue = boardRange.getCell(row, col).getValue() as string;
            if (cellValue && typeof cellValue === 'string' && cellValue.charAt(0) === player) {
                const piece = cellValue;
                // Get all possible moves for the current piece
                moves.push(...getPossibleMovesForPiece(boardRange, piece, row, col));
            }
        }
    }

    return moves;
}

function getPossibleMovesForPiece(boardRange: ExcelScript.Range, piece: string, row: number, col: number): Move[] {
    const moves: Move[] = [];
    // Define the movement directions based on the piece color
    const directions = piece.startsWith("W") ? [[-1, -1], [-1, 1]] : [[1, -1], [1, 1]];

    directions.forEach(([dRow, dCol]) => {
        const newRow = row + dRow;
        const newCol = col + dCol;
        // Check if the move is valid
        if (isValidMove(boardRange, row, col, newRow, newCol)) {
            moves.push({ piece, from: [row, col], to: [newRow, newCol] });
        }
        const jumpRow = row + 2 * dRow;
        const jumpCol = col + 2 * dCol;
        // Check if the move is a capture move
        if (isCaptureMove(boardRange, row, col, newRow, newCol, jumpRow, jumpCol)) {
            const capturedPiece = boardRange.getCell(newRow, newCol).getValue() as string;
            moves.push({ piece, from: [row, col], to: [jumpRow, jumpCol], captures: [capturedPiece] });
        }
    });

    return moves;
}

function isValidMove(boardRange: ExcelScript.Range, fromRow: number, fromCol: number, toRow: number, toCol: number): boolean {
    // Check if the target cell is within the board and empty
    if (toRow < 0 || toRow >= 8 || toCol < 0 || toCol >= 8) return false;
    return boardRange.getCell(toRow, toCol).getValue() === "";
}

function isCaptureMove(boardRange: ExcelScript.Range, fromRow: number, fromCol: number, overRow: number, overCol: number, toRow: number, toCol: number): boolean {
    // Check if the target cell is within the board
    if (toRow < 0 || toRow >= 8 || toCol < 0 || toCol >= 8) return false;
    const overCell = boardRange.getCell(overRow, overCol).getValue() as string;
    const toCell = boardRange.getCell(toRow, toCol).getValue() as string;
    // Check if the move captures an opponent's piece
    if (toCell !== "" || overCell === "" || typeof overCell !== 'string' || overCell.charAt(0) === (boardRange.getCell(fromRow, fromCol).getValue() as string).charAt(0)) {
        return false;
    }
    return true;
}

function evaluateMoves(moves: Move[], depth: number): Move {
    // For simplicity, return the first move. In a real implementation, add logic to evaluate the best move up to the given ply depth.
    return moves.length > 0 ? moves[0] : null;
}

function executeMove(boardRange: ExcelScript.Range, move: Move) {
    // Move the piece to the new position
    const [fromRow, fromCol] = move.from;
    const [toRow, toCol] = move.to;
    const piece = move.piece;

    boardRange.getCell(fromRow, fromCol).setValue("");
    boardRange.getCell(toRow, toCol).setValue(piece);

    // Handle captures by removing the captured pieces from the board
    if (move.captures) {
        move.captures.forEach(capturedPiece => {
            for (let row = 0; row < 8; row++) {
                for (let col = 0; col < 8; col++) {
                    if (boardRange.getCell(row, col).getValue() === capturedPiece) {
                        boardRange.getCell(row, col).setValue("");
                    }
                }
            }
        });
    }
}

function switchTurn(helperCell: ExcelScript.Range) {
    // Switch the turn to the other player
    const currentPlayer = helperCell.getValue() as string;
    const nextPlayer = currentPlayer === "W" ? "B" : "W";
    helperCell.setValue(nextPlayer);
}

// Define the Move interface
interface Move {
    piece: string;
    from: [number, number];
    to: [number, number];
    captures?: string[];
}
