# VBA Conway Game of Life

An object-oriented implementation of **Conway’s Game of Life** in **VBA for Microsoft Excel**.  
The project demonstrates modular design in VBA, separating simulation logic, rendering, and configuration into dedicated classes and interfaces.

## Demo Video
Watch the simulation in action on [YouTube](https://youtu.be/4DAZ3WulKko?feature=shared).

---

## Features
- Classic Conway’s Game of Life rules  
- Modular class structure:
  - `clsGameOfLife` — automaton rules
  - `clsGridRenderer` — renders the grid in Excel
  - `clsSimulation` — manages the simulation loop
  - `clsSimulationConfig` — stores grid size, wrapping, and options
  - `ICellularAutomaton` — interface for cellular automata
  - `IRenderer` — interface for renderers
  - `mdMain` — entry point to run the simulation
- Configurable grid size and wrapping  
- Easily extensible with new automata or renderers  
- Clean, consistent naming conventions  

---

## Installation

You have two options:

### 1. Quick Start (recommended)
Download the provided **`ConwaysGameOfLifeDemo.xlsm`** workbook, open it in Excel, enable macros, and run the simulation.

### 2. Manual Setup
1. Download or clone this repository.  
2. Open Excel and press `Alt + F11` to open the VBA editor.  
3. Import the `.cls` and `.bas` files into your VBA project:  
   - `File > Import File…` for each file.  
4. Save the workbook as a **macro-enabled workbook** (`.xlsm`).  
5. Ensure macros are enabled in Excel.  

---

## Usage
1. Open the VBA editor (`Alt + F11`).  
2. Run the simulation entry point:  
   - In the editor, select `mdMain.StartSimulation` and press `F5`.  
   - Or assign `StartSimulation` to a button on your worksheet.  
3. The Game of Life will render on the worksheet and evolve step by step.  

Configuration (grid size, wrapping, etc.) can be adjusted in `clsSimulationConfig`.

---

## Contributing
Contributions are welcome! Ideas include:
- New rendering styles (colors, conditional formatting, shapes)  
- Alternative cellular automata rules (HighLife, Brian’s Brain, etc.)  
- Performance improvements for larger grids  

---

## License
This project is released under the [MIT License](LICENSE).
