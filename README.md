# Snake Game in Excel (VBA)

A fully playable Snake game built inside an Excel spreadsheet. Every square
on the board is a cell, and VBA macros move the snake, place the food and
keep score.

![Gameplay screenshot](screenshot.png)

## Demo
Watch the 20-second demo video: add your video link here

## Features
- 20 x 20 game board made from Excel cells
- Steer with the arrow keys or W, A, S, D
- Eat the amber food to grow longer and score 10 points
- Game over when the snake hits a wall or its own body
- High score saved in a cell
- START, PAUSE and RESET buttons
- Live status box (PLAYING, PAUSED, GAME OVER, NEW HIGH SCORE!)

## How to play
1. Download `Snake_Game.xlsm`.
2. Open it in desktop Excel for Windows and click Enable Content.
3. Click START.
4. Steer with the arrow keys or W, A, S, D, and eat the food.
5. Click PAUSE to pause or resume, and RESET to clear the board.

Note: the game needs desktop Excel on Windows. It will not run in Excel
Online, Google Sheets, on a phone or on a Mac.

## How it works
- The board is a named range (Board) of 20 x 20 cells. Colouring a cell
  draws the snake, the food or the empty space.
- The snake's position is stored in two arrays (columns and rows). Each
  step, every body segment moves into the place of the one in front of it
  and the head moves one square in the current direction.
- A loop with a short pause (0.2 seconds) moves the snake automatically.
  A smaller pause number makes the game faster.
- Key presses are read with the Windows function GetAsyncKeyState, so
  steering responds instantly. The code also blocks reversing straight
  into the snake's own body.
- Food is placed on a random empty square.
- Before each move, the code checks the next square for a wall or a body
  segment to decide whether the game is over.

## Skills used
VBA (macros, loops, arrays, functions), named ranges, Windows API call,
game logic, cell formatting, form-control buttons, dashboard-style layout

## Files
- Snake_Game.xlsm: the playable game
- Snake.bas: the exported VBA code, so you can read it without opening Excel
- screenshot.png and the demo video

## Author
Built by Basit Jimoh
