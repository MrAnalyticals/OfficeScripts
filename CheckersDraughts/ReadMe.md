**Checkers**, also known as Draughts in some countries, can , now, be played within Excel Online. The code is less than 160 lines long. 
Copy the code into your own tenant and try it out. The code, also, builds your playing board for you! 

![image](https://github.com/user-attachments/assets/3be85a2e-ba0b-41b7-af85-f5b42f0dbbb9)

The two YouTube Videos for this are here: 

Demo 1: https://www.youtube.com/watch?v=5TAmScVYGt0
![image](https://github.com/user-attachments/assets/e7788296-ca44-41ab-adea-bf814e7c7117)



Demo 2: 

**Video 2 Dialog**

Following on from my introductory video where I described how it was possible to build a game  of draughts to work in Excel online using office scripts, let’s dive straight into the code 
This checkers game uses the Minimax algorithm for decision-making. 
Let’s first describe the code logic:
1. The game starts by initializing the board and setting the first player to White.
2. For each turn, the code finds all possible moves, evaluates them using Minimax, and chooses the best one.
3. The best move is executed, pieces are moved, and the turn is passed to the opponent.

Let review each of the functions. 
The Main Function:
   - It starts the game, initializes the board if necessary, and identifies the current player. 
   - Then, it evaluates all possible moves using the evaluate Moves function, which implements the Minimax algorithm up to 4 plies depth).
   - After choosing the best move, it executes the move and switches turns between players.

The interface is a return type. It defines a move in the game, including the piece being moved, its starting and ending positions, and any captured pieces during the move.

The evaluate Moves Function:
   - This function takes the possible moves, the board state, and a depth parameter.
   - The Minimax algorithm is used here to explore different game states. 
   - It simulates several future moves, alternating between players (maximizing for one player, minimizing for the other), and evaluates which move leads to the most favorable outcome.

The Minimax Algorithm:
   - The Minimax alternates between maximizing and minimizing players:
     - The maximizing player (usually White) tries to maximize the board score by making advantageous moves.
     - The minimizing player (usually Black) tries to minimize the score by making moves that counter the opponent.
   - The code looks at the game tree, simulates future moves, and at a certain depth (4 plies in this case), evaluates the board using a simple heuristic: each player’s pieces are worth points.
   - the minimax function recursively calls itself until it reaches the desired depth or when no more moves are possible. It returns the score of the board and helps decide the best move at each level of the game tree.

The  Heuristic Evaluation evaluateBoard:
   - This function assigns scores to the board state. White pieces add points, Black pieces subtract points, providing a simple way to judge the state of the game.


