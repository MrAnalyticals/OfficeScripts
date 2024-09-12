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

    // Evaluate the possible moves up to 4 plies deep using minimax with alpha-beta pruning
    const bestMove = evaluateMoves(boardRange, possibleMoves, currentPlayer, 4, Number.NEGATIVE_INFINITY, Number.POSITIVE_INFINITY, true);

    // Execute the best move if one is found
    if (bestMove) {
        executeMove(boardRange, bestMove);
        // Switch turns after executing the move
        switchTurn(helperCell);
    }
}

function evaluateMoves(boardRange: ExcelScript.Range, moves: Move[], currentPlayer: string, depth: number, alpha: number, beta: number, maximizingPlayer: boolean): Move {
    if (depth === 0 || moves.length === 0) {
        return null; // Stop evaluation if depth limit is reached or no moves are left
    }

    let bestMove: Move = null;
    let bestScore = maximizingPlayer ? Number.NEGATIVE_INFINITY : Number.POSITIVE_INFINITY;

    for (let move of moves) {
        // Simulate the move
        const simulatedBoard = simulateMove(boardRange, move);

        // Get the opponent's possible moves after this move
        const opponent = currentPlayer === "W" ? "B" : "W";
        const opponentMoves = getAllPossibleMoves(simulatedBoard, opponent);

        // Recursively evaluate the opponent's moves
        const childMove = evaluateMoves(simulatedBoard, opponentMoves, opponent, depth - 1, alpha, beta, !maximizingPlayer);

        // Simple score assignment (can be enhanced with more sophisticated evaluation criteria)
        const moveScore = childMove ? (maximizingPlayer ? 1 : -1) : 0;

        if (maximizingPlayer) {
            if (moveScore > bestScore) {
                bestScore = moveScore;
                bestMove = move;
            }
            alpha = Math.max(alpha, moveScore);
        } else {
            if (moveScore < bestScore) {
                bestScore = moveScore;
                bestMove = move;
            }
            beta = Math.min(beta, moveScore);
        }

        // Alpha-beta pruning
        if (beta <= alpha) {
            break;
        }
    }

    return bestMove;
}

function simulateMove(boardRange: ExcelScript.Range, move: Move): ExcelScript.Range {
    // Create a copy of the board
    const simulatedBoard = boardRange.getWorksheet().getRange("A1:H8");
    
    // Simulate the move on the copied board
    const [fromRow, fromCol] = move.from;
    const [toRow, toCol] = move.to;
    const piece = move.piece;

    simulatedBoard.getCell(fromRow, fromCol).setValue("");
    simulatedBoard.getCell(toRow, toCol).setValue(piece);

    // Handle captures by removing the captured pieces from the board
    if (move.captures) {
        move.captures.forEach(capturedPiece => {
            for (let row = 0; row < 8; row++) {
                for (let col = 0; col < 8; col++) {
                    if (simulatedBoard.getCell(row, col).getValue() === capturedPiece) {
                        simulatedBoard.getCell(row, col).setValue("");
                    }
                }
            }
        });
    }

    return simulatedBoard;
}
