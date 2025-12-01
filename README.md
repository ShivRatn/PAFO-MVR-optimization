#  Optimization_PAFO-MVR.py: Hybrid Differential Evolution
This Python script implements a specialized optimization algorithm designed to find the **global minimum of output 'B'** within a **tightly constrained feasible region of output 'A'** by manipulating two input variables (X1 and X2) in an external Microsoft Excel model.

The core challenge is the existence of **multiple, small, and distinct "feasible regions"** within the search space, which requires a robust multi-modal global search strategy. The script addresses this using an **Enhanced Differential Evolution (DE)** approach.

---

##  Prerequisites

To run this script successfully, the following are required:

1.  **Python Environment:** Python 3.x
2.  **Required Libraries:**
    ```bash
    pip install xlwings numpy scipy
    ```
3.  **Microsoft Excel:** The script requires a working installation of Microsoft Excel to communicate with the workbook.
4.  **Excel Workbook:** The file `v1.xlsx` must be present in the same directory as the script.
5.  **Workbook Configuration:** The Excel file must contain a sheet named `Fertigation` with a formula-based model where:
    * **Inputs (X1, X2)** are read from cells **L4** and **L5**.
    * **Output A** (Constraint/Feasibility check) is in cell **D26**.
    * **Output B** (Objective to Minimize) is in cell **O11**.

---

##  Optimization Goal & Constraints

### Goal
Minimize **Output B** (`O11`)

### Tight Feasibility Constraint
**Output A** (`D26`) must be within the extremely narrow range:
$$25.000 \leq A \leq 25.100$$
(This is controlled by `feasibility_tolerance=0.1` around a target of 25.05)

### Input Constraint
The two input variables, **X1 (L4)** and **X2 (L5)**, must be **exact multiples of 10**.
* **Search Bounds (Scaled):**
    * $X1$: $[3500, 8500]$
    * $X2$: $[4000, 12000]$
    * *(Note: The script internally scales the bounds down by a factor of 10 for optimization, meaning internal variables are in the range of, for example, [350, 850], and are snapped to a grid step of 1.0 (which corresponds to 10 in the actual Excel input).*

---

##  Technical Flow and Strategy

The `SingleStrategyFeasibilityOptimizer` class orchestrates the process in three main phases, all leveraging the robust and multi-modal exploration capabilities of the **Differential Evolution (DE)** algorithm.

### 1. Excel Setup and Management
* **`setup_excel()`:** Initializes the `xlwings` application in a **headless, non-interactive mode** (`visible=False`, `screen_updating=False`, `calculation='manual'`) for maximum performance and stability during repeated calculations.
* **`evaluate_excel(params)`:** The core function for model evaluation. It ensures:
    1.  Inputs (`params`) are **snapped to the required grid** (multiples of 10 in the actual input space).
    2.  A **caching mechanism** (`evaluation_cache`) is used to avoid redundant Excel calls for identical inputs.
    3.  Inputs are written to Excel, the workbook is calculated (`self.sheet.book.app.calculate()`), and outputs A and B are read back.

### 2. PHASE 1: Enhanced Global Exploration (65% Budget)

This phase uses multiple, diverse runs of the **Differential Evolution** algorithm to maximize the chance of discovering all small, distant feasible regions.

* **Adaptive Objective Function:** The objective function in this phase is specifically designed to reward **Feasibility** over simply minimizing B.
    * **Feasible Points:** Return a large negative value (e.g., `-1000`) plus a small, random reward/penalty to encourage diverse movement and exploration (`np.random.random()`). Different DE runs use varying rewards (e.g., one rewards distance from existing regions; another slightly favors better B values) to enhance diversity.
    * **Infeasible Points:** Return a **quadratic penalty** based on the distance from the constraint window $[25.000, 25.100]$.
* **Region Tracking:** The `find_or_update_region` method tracks discovered feasible points, clustering them into distinct `FeasibleRegion` objects if they are farther than `min_region_distance` (default: 20 units, or 200 in actual input space) from existing regions.

### 3. PHASE 2: Local Exploitation (35% Budget)

Once feasible regions are found, the focus shifts to finding the true minimum of B *within* each region.

* **Local Bounds:** A smaller, focused search space is defined around the center of each discovered region.
* **Optimization Strategies:** The budget for each region is split:
    1.  **Fine-grained DE (70%):** A local, low-population DE run focused on finding the optimum within the region. The objective function here prioritizes **minimizing B** (`return y2`) and applies a **heavy penalty** (multiplied by 1000) for points that fall *outside* the tight feasibility constraint.
    2.  **Pattern Search (30%):** A custom coordinate descent/pattern search is performed using decreasing, grid-snapped step sizes (e.g., 5.0, 2.0, 1.0) to ensure the final solution is precisely on the required **multiples of 10** grid. This refines the best solution found by DE.

### 4. Finalization
* The script compares the `best_b_value` found in all `FeasibleRegion` objects to determine the **Global Optimum**.
* The final optimal point is **re-evaluated outside the cache** (manual verification) to confirm its A and B values and adherence to the multiples of 10 constraint.

---

##  Running the Script

1.  Place `Optimization_PAFO-MVR.py` and `v1.xlsx` in the same directory.
2.  Open a terminal or command prompt in that directory.
3.  Run the script:

    ```bash
    python Optimization_PAFO-MVR.py
    ```

### Expected Output

The script will print detailed progress reports for the Excel setup, connection testing, Phase 1 (DE runs), and Phase 2 (Region Optimization). The final output will clearly state the global optimum:
