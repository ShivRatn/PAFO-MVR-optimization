import xlwings as xw
import numpy as np
import time
from scipy.optimize import differential_evolution, minimize
import random
from dataclasses import dataclass
from typing import List, Tuple, Optional
import copy

# === USER SETTINGS ===
excel_file = "v1.xlsx"
sheet_name = "Fertigation"

input_cell_1 = "L4"
input_cell_2 = "L5"
output_cell_1 = "D26"  # Output 'A'
output_cell_2 = "O11"  # Output 'B'

# Search bounds (scaled down by 10 for internal use)
bounds = [(3500, 8500), (4000, 12000)]

@dataclass
class FeasibleRegion:
    """Represents a discovered feasible region"""
    center: List[float]
    best_solution: List[float]
    best_b_value: float
    best_a_value: float
    radius: float
    evaluations_count: int
    fully_explored: bool = False

@dataclass
class OptimizationResult:
    best_x1: float
    best_x2: float
    best_a: float
    best_b: float
    total_evaluations: int
    feasible_regions_found: int
    success: bool
    algorithm_used: str

class SingleStrategyFeasibilityOptimizer:
    """
    Simplified optimizer using only Differential Evolution for multi-modal feasibility problems
    
    Strategy:
    1. ENHANCED GLOBAL EXPLORATION: Single DE strategy with maximum diversity
    2. LOCAL EXPLOITATION: Thorough search within each found region  
    3. REGION COMPARISON: Return globally optimal solution across all regions
    
    CONSTRAINT: All input values must be multiples of 10
    """
    
    def __init__(self, feasibility_tolerance=0.1, min_region_distance=200):
        self.sheet = None
        self.eval_count = 0
        self.evaluation_cache = {}
        
        # Feasible region tracking
        self.feasible_regions: List[FeasibleRegion] = []
        self.feasibility_tolerance = feasibility_tolerance  # A constraint window: 25.0 to 25.1
        self.min_region_distance = min_region_distance  # Minimum distance between regions
        
        # Search configuration - More budget for single DE strategy
        self.global_exploration_budget = 0.65  # 65% budget for finding regions (increased)
        self.local_exploitation_budget = 0.35  # 35% budget for refining regions
        
    def round_to_step(self, value, step=1.0):
        """Round value to nearest multiple of step (step=1.0 means multiples of 10 in actual inputs)"""
        return round(value / step) * step
    
    def snap_params_to_grid(self, params):
        """Snap parameters to multiples of 10 in actual input space (step=1 in scaled space)"""
        return [self.round_to_step(params[0], 1.0), self.round_to_step(params[1], 1.0)]
        
    def setup_excel(self):
        """Setup Excel for optimal performance"""
        try:
            self.app = xw.App(visible=False)
            self.app.screen_updating = False
            self.app.calculation = 'manual'
            self.app.display_alerts = False
            
            wb = xw.Book(excel_file)
            self.sheet = wb.sheets[sheet_name]
            
            # Test Excel connection
            print(f" Excel setup complete. Testing connection...")
            test_result = self.sheet.range("A1").value  # Test read
            print(f" Excel connection verified")
            
        except Exception as e:
            print(f" Excel setup failed: {str(e)}")
            raise
        
    def cleanup_excel(self):
        """Restore Excel and close"""
        if hasattr(self, 'app') and self.app is not None:
            try:
                self.app.screen_updating = True
                self.app.calculation = 'automatic'  
                self.app.display_alerts = True
                self.app.quit()
            except Exception as e:
                print(f" Excel cleanup warning: {str(e)}")
        
    def test_excel_connection(self):
        """Test Excel connection and verify a known input/output pair"""
        print(f"\n TESTING EXCEL CONNECTION AND I/O VERIFICATION")
        
        # Test with a simple known case
        test_x1, test_x2 = 70000, 85000  # Multiples of 10
        
        try:
            self.sheet.range(input_cell_1).value = test_x1
            self.sheet.range(input_cell_2).value = test_x2
            
            # Try calculation
            self.sheet.book.app.calculate()
            time.sleep(0.02)
            
            test_y1 = self.sheet.range(output_cell_1).value
            test_y2 = self.sheet.range(output_cell_2).value
            
            print(f" Test evaluation successful:")
            print(f"   Input: X1={test_x1}, X2={test_x2}")
            print(f"   Output: A={test_y1}, B={test_y2}")
            print(f"   Feasible: {self.is_feasible(test_y1)}")
            
            return True
            
        except Exception as e:
            print(f" Excel connection test failed: {str(e)}")
            return False
    
    def evaluate_excel(self, params):
        """Fast Excel evaluation with caching - ensures inputs are multiples of 10"""
        # Snap to grid first (ensures multiples of 10 in actual input space)
        snapped_params = self.snap_params_to_grid(params)
        x1, x2 = snapped_params[0] * 10, snapped_params[1] * 10
        
        # Ensure inputs are exactly multiples of 10
        x1 = round(x1 / 10) * 10
        x2 = round(x2 / 10) * 10
        
        # Check cache (using snapped values)
        cache_key = (int(x1), int(x2))
        if cache_key in self.evaluation_cache:
            return self.evaluation_cache[cache_key]
        
        try:
            # Excel evaluation with more robust calculation method
            self.sheet.range(input_cell_1).value = x1
            self.sheet.range(input_cell_2).value = x2
            
            # Try multiple calculation methods in order of preference
            try:
                # Method 1: Standard calculate
                self.sheet.book.app.calculate()
            except:
                try:
                    # Method 2: Full calculation
                    self.app.calculation = 'automatic'
                    self.app.calculation = 'manual'
                except:
                    # Method 3: Range-specific calculation
                    self.sheet.calculate()
            
            # Small delay to ensure calculation completes
            time.sleep(0.01)
            
            y1 = self.sheet.range(output_cell_1).value
            y2 = self.sheet.range(output_cell_2).value
            
            self.eval_count += 1
            
            # Validate outputs
            if y1 is None or y2 is None or isinstance(y1, str) or isinstance(y2, str):
                result = (None, None)
            else:
                result = (float(y1), float(y2))
            
            # Cache and return
            self.evaluation_cache[cache_key] = result
            
            return result
            
        except Exception as e:
            print(f"    Excel evaluation error at X1={x1}, X2={x2}: {str(e)}")
            return (None, None)

    def is_feasible(self, a_value):
        """Check if solution satisfies the tight feasibility constraint"""
        return a_value is not None and 25.000 <= a_value <= 25.1

    def euclidean_distance(self, point1, point2):
        """Calculate distance between two points"""
        return np.sqrt(sum((p1 - p2)**2 for p1, p2 in zip(point1, point2)))

    def find_or_update_region(self, params, a_value, b_value):
        """Add point to existing region or create new region"""
        # Snap params to grid before processing
        snapped_params = self.snap_params_to_grid(params)
        
        # Check if this point belongs to an existing region
        for region in self.feasible_regions:
            if self.euclidean_distance(snapped_params, region.center) <= self.min_region_distance:
                # Update existing region
                region.evaluations_count += 1
                if b_value < region.best_b_value:
                    region.best_solution = snapped_params[:]
                    region.best_b_value = b_value
                    region.best_a_value = a_value
                    print(f" Updated Region {len(self.feasible_regions)}: New best B={b_value:.4f}")
                return region
        
        # Create new region
        new_region = FeasibleRegion(
            center=snapped_params[:],
            best_solution=snapped_params[:],
            best_b_value=b_value,
            best_a_value=a_value,
            radius=50.0,  # Initial search radius
            evaluations_count=1
        )
        self.feasible_regions.append(new_region)
        print(f" NEW FEASIBLE REGION #{len(self.feasible_regions)} FOUND!")
        print(f"   Location: ({snapped_params[0]*10:.0f}, {snapped_params[1]*10:.0f})")
        print(f"   A={a_value:.4f}, B={b_value:.4f}")
        return new_region

    # =============================================================================
    # PHASE 1: ENHANCED GLOBAL EXPLORATION - Single DE Strategy Only
    # =============================================================================
    
    def enhanced_global_exploration(self, budget):
        """
        Single but enhanced Differential Evolution strategy for comprehensive region discovery
        Uses multiple DE runs with different configurations for maximum diversity
        """
        print(f"\n PHASE 1: ENHANCED GLOBAL EXPLORATION ({budget} evaluations)")
        print(" Goal: Find all tiny feasible regions using enhanced DE strategy")
        print(" Constraint: All inputs must be multiples of 10")
        
        # Split budget across multiple DE runs with different configurations
        num_de_runs = 4  # Multiple DE runs for diversity
        budget_per_run = budget // num_de_runs
        
        de_configs = [
            # Config 1: High diversity, large population
            {"popsize": 15, "mutation": (0.5, 1.5), "recombination": 0.9, "focus": "maximum_diversity"},
            # Config 2: Medium diversity, adaptive
            {"popsize": 10, "mutation": (0.3, 1.2), "recombination": 0.7, "focus": "adaptive_search"}, 
            # Config 3: Focused exploration around found regions
            {"popsize": 8, "mutation": (0.4, 1.0), "recombination": 0.8, "focus": "region_expansion"},
            # Config 4: Fine-grained search
            {"popsize": 12, "mutation": (0.2, 0.9), "recombination": 0.6, "focus": "fine_grained"}
        ]
        
        for run_idx, config in enumerate(de_configs):
            print(f"\n DE Run {run_idx + 1}/4: {config['focus']} ({budget_per_run} evals)")
            self.run_diverse_differential_evolution(budget_per_run, config, run_idx)
        
        print(f"\n ENHANCED EXPLORATION RESULTS:")
        print(f"    Feasible regions found: {len(self.feasible_regions)}")
        print(f"    Total evaluations used: {self.eval_count}")
        
        return len(self.feasible_regions) > 0

    def run_diverse_differential_evolution(self, budget, config, run_idx):
        """Enhanced DE run with specific configuration"""
        
        def enhanced_exploration_objective(params):
            """Enhanced objective function that adapts based on run configuration and findings"""
            y1, y2 = self.evaluate_excel(params)
            
            if y1 is None or y2 is None:
                return 1e6
                
            if self.is_feasible(y1):
                # Found feasible point - record it
                self.find_or_update_region(params, y1, y2)
                
                # Adaptive reward based on run configuration
                base_reward = -1000
                
                if config['focus'] == 'maximum_diversity':
                    # Maximum randomness for diversity
                    return base_reward - np.random.random() * 100
                elif config['focus'] == 'adaptive_search':
                    # Slight preference for better B values but still exploratory
                    return base_reward - y2 * 0.1 - np.random.random() * 50
                elif config['focus'] == 'region_expansion':
                    # Encourage exploration around existing regions
                    if len(self.feasible_regions) > 1:
                        # Small bonus for points far from existing regions
                        min_dist = min(self.euclidean_distance(params, r.center) 
                                     for r in self.feasible_regions)
                        distance_bonus = min(min_dist * 0.1, 50)
                        return base_reward - distance_bonus - np.random.random() * 30
                    else:
                        return base_reward - np.random.random() * 50
                else:  # fine_grained
                    # Balance between exploration and exploitation
                    return base_reward - y2 * 0.2 - np.random.random() * 25
                    
            else:
                # Infeasible - guide toward feasible regions with enhanced penalties
                if y1 < 25.000:
                    penalty = (25.000 - y1) ** 2
                else:
                    penalty = (y1 - 25.1) ** 2
                
                # Add noise to avoid getting stuck in local minima
                noise = np.random.random() * 10
                return penalty + noise
        
        # Adaptive bounds - expand search in later runs if regions found
        search_bounds = bounds
        if run_idx >= 2 and len(self.feasible_regions) > 0:
            # Expand search around found regions for runs 3 and 4
            all_centers = [r.center for r in self.feasible_regions]
            min_x1 = min(c[0] for c in all_centers) - 300
            max_x1 = max(c[0] for c in all_centers) + 300
            min_x2 = min(c[1] for c in all_centers) - 300
            max_x2 = max(c[1] for c in all_centers) + 300
            
            search_bounds = [
                (max(bounds[0][0], min_x1), min(bounds[0][1], max_x1)),
                (max(bounds[1][0], min_x2), min(bounds[1][1], max_x2))
            ]
            print(f"   🎯 Focused search bounds: {search_bounds}")
        
        # Run DE with configuration
        try:
            result = differential_evolution(
                enhanced_exploration_objective,
                search_bounds,
                maxiter=budget // config['popsize'],
                popsize=config['popsize'],
                seed=None,  # No seed for maximum randomness
                mutation=config['mutation'],
                recombination=config['recombination'],
                atol=1e-12,  # Don't converge early
                tol=1e-12,
                polish=False,
                updating='immediate'
            )
            
            print(f"    DE Run {run_idx + 1} complete. Regions found so far: {len(self.feasible_regions)}")
            
        except Exception as e:
            print(f"    DE Run {run_idx + 1} encountered error: {str(e)}")
            # Continue with remaining runs

    # =============================================================================
    # PHASE 2: LOCAL EXPLOITATION - Same as before
    # =============================================================================

    def local_exploitation_phase(self, budget):
        """
        Intensive local search within each discovered feasible region
        Goal: Find the true optimum (minimum B) within each region
        """
        if not self.feasible_regions:
            print(" No feasible regions found - skipping local exploitation")
            return
            
        print(f"\n PHASE 2: LOCAL EXPLOITATION ({budget} evaluations)")
        print(f" Goal: Find minimum B within each of {len(self.feasible_regions)} regions")
        print(" Constraint: All inputs must be multiples of 10")
        
        budget_per_region = budget // len(self.feasible_regions)
        
        for region_idx, region in enumerate(self.feasible_regions):
            print(f"\n Optimizing Region #{region_idx + 1}")
            print(f"   Center: ({region.center[0]*10:.0f}, {region.center[1]*10:.0f})")
            print(f"   Current best B: {region.best_b_value:.4f}")
            
            self.optimize_within_region(region, budget_per_region)

    def optimize_within_region(self, region: FeasibleRegion, budget):
        """Intensive optimization within a single feasible region"""
        
        # Define local bounds around the region
        expansion_factor = 1.5  # Search slightly beyond known region
        local_bounds = [
            (max(bounds[0][0], region.center[0] - region.radius * expansion_factor),
             min(bounds[0][1], region.center[0] + region.radius * expansion_factor)),
            (max(bounds[1][0], region.center[1] - region.radius * expansion_factor),
             min(bounds[1][1], region.center[1] + region.radius * expansion_factor))
        ]
        
        def local_objective(params):
            """Objective for local optimization within region"""
            y1, y2 = self.evaluate_excel(params)
            
            if y1 is None or y2 is None:
                return 1e6
                
            if self.is_feasible(y1):
                # Update region if better solution found
                snapped_params = self.snap_params_to_grid(params)
                if y2 < region.best_b_value:
                    region.best_solution = snapped_params[:]
                    region.best_b_value = y2
                    region.best_a_value = y1
                return y2  # Minimize B
            else:
                # Heavy penalty for leaving feasible region
                if y1 < 25.000:
                    penalty = (25.000 - y1) ** 2 * 1000
                else:
                    penalty = (y1 - 25.1) ** 2 * 1000
                return penalty + 500

        # Strategy 1: Fine-grained DE within region (70% of budget)
        de_budget = int(budget * 0.7)
        if de_budget > 20:
            print(f"    Local DE optimization ({de_budget} evals)")
            differential_evolution(
                local_objective,
                local_bounds,
                maxiter=de_budget // 20,
                popsize=5,  # Small population for local search
                seed=42,
                mutation=(0.2, 0.8),  # Lower mutation for local search
                recombination=0.7,
                polish=False
            )
        
        # Strategy 2: Gradient-free local search (30% of budget)
        remaining_budget = budget - de_budget
        if remaining_budget > 10:
            print(f"    Pattern search optimization ({remaining_budget} evals)")
            self.pattern_search_within_region(region, local_bounds, remaining_budget)
        
        print(f"    Region optimization complete. Best B: {region.best_b_value:.4f}")

    def pattern_search_within_region(self, region: FeasibleRegion, local_bounds, budget):
        """Coordinate descent / pattern search within region - multiples of 10 only"""
        current_point = region.best_solution[:]
        # Step sizes in scaled space (1.0 = 10 in actual space, 5.0 = 50 in actual space, etc.)
        step_sizes = [5.0, 2.0, 1.0]  # Corresponds to steps of 50, 20, 10 in actual input space
        
        for step_size in step_sizes:
            if budget <= 0:
                break
                
            print(f"     Pattern search with step size {step_size} (={step_size*10:.0f} in actual input)")
            improved = True
            while improved and budget > 0:
                improved = False
                
                # Try all coordinate directions
                for dim in range(len(current_point)):
                    if budget <= 0:
                        break
                        
                    # Try positive and negative steps
                    for direction in [1, -1]:
                        if budget <= 0:
                            break
                            
                        # Create neighbor
                        neighbor = current_point[:]
                        neighbor[dim] += direction * step_size
                        
                        # Ensure within bounds and snap to grid
                        neighbor[dim] = np.clip(neighbor[dim], 
                                              local_bounds[dim][0], 
                                              local_bounds[dim][1])
                        neighbor = self.snap_params_to_grid(neighbor)
                        
                        # Skip if same as current point (due to snapping/bounds)
                        if neighbor == current_point:
                            continue
                        
                        # Evaluate
                        y1, y2 = self.evaluate_excel(neighbor)
                        budget -= 1
                        
                        if (y1 is not None and y2 is not None and 
                            self.is_feasible(y1) and y2 < region.best_b_value):
                            
                            # Better solution found
                            region.best_solution = neighbor[:]
                            region.best_b_value = y2
                            region.best_a_value = y1
                            current_point = neighbor[:]
                            improved = True
                            print(f"        Improved: B={y2:.4f} at ({neighbor[0]*10:.0f}, {neighbor[1]*10:.0f})")
                            break  # Move to next dimension
        
        region.fully_explored = True

    # =============================================================================
    # MAIN OPTIMIZATION ORCHESTRATOR
    # =============================================================================

    def optimize(self, max_evaluations=10000):
        """
        Main optimization orchestrator using single enhanced DE strategy
        """
        self.setup_excel()
        
        try:
            print(f" SINGLE-STRATEGY FEASIBILITY OPTIMIZER")
            print(f" Total Budget: {max_evaluations} evaluations")
            print(f" Problem: Multiple tiny feasible regions (A ∈ [25.000, 25.100])")
            print(f" Goal: Find global minimum of B across all feasible regions")
            print(f" Strategy: Enhanced Differential Evolution Only")
            print(f" CONSTRAINT: All input values must be multiples of 10")
            
            # Test Excel connection first
            if not self.test_excel_connection():
                return OptimizationResult(0, 0, 0, 0, 0, 0, False, "Excel Connection Failed")
            
            start_time = time.time()
            
            # Phase 1: Enhanced global exploration using only DE
            exploration_budget = int(max_evaluations * self.global_exploration_budget)
            regions_found = self.enhanced_global_exploration(exploration_budget)
            
            if not regions_found:
                print("\n OPTIMIZATION FAILED: No feasible regions discovered")
                return OptimizationResult(0, 0, 0, 0, self.eval_count, 0, False, "Single-Strategy Optimizer")
            
            # Phase 2: Optimize within each region
            exploitation_budget = max_evaluations - self.eval_count
            if exploitation_budget > 0:
                self.local_exploitation_phase(exploitation_budget)
            
            # Phase 3: Select global optimum
            return self.finalize_optimization(start_time)
            
        except Exception as e:
            print(f" Optimization failed with error: {str(e)}")
            return OptimizationResult(0, 0, 0, 0, self.eval_count, 0, False, f"Error: {str(e)}")
        finally:
            self.cleanup_excel()

    def finalize_optimization(self, start_time):
        """Select the globally optimal solution across all regions"""
        end_time = time.time()
        
        print(f"\n{'='*60}")
        print(" SINGLE-STRATEGY OPTIMIZATION COMPLETE")
        print(f"{'='*60}")
        print(f" Total time: {end_time - start_time:.1f} seconds")
        print(f" Total evaluations: {self.eval_count}")
        print(f" Feasible regions found: {len(self.feasible_regions)}")
        print(f" Cache efficiency: {len(self.evaluation_cache)}/{self.eval_count} = {len(self.evaluation_cache)/max(self.eval_count,1)*100:.1f}%")
        print(f" Strategy: Enhanced Differential Evolution Only")
        print(f" Input constraint: All values are multiples of 10")
        
        if not self.feasible_regions:
            return OptimizationResult(0, 0, 0, 0, self.eval_count, 0, False, "Single-Strategy Optimizer")
        
        # Find globally optimal region
        global_best_region = min(self.feasible_regions, key=lambda r: r.best_b_value)
        
        # VERIFICATION: Re-evaluate the best solution to confirm results
        print(f"\n VERIFICATION: Re-evaluating optimal solution...")
        verification_x1 = global_best_region.best_solution[0] * 10
        verification_x2 = global_best_region.best_solution[1] * 10
        
        # Ensure inputs are multiples of 10
        verification_x1 = round(verification_x1 / 10) * 10
        verification_x2 = round(verification_x2 / 10) * 10
        
        # Manual verification (bypass cache)
        self.sheet.range(input_cell_1).value = verification_x1
        self.sheet.range(input_cell_2).value = verification_x2
        try:
            self.sheet.book.app.calculate()
            time.sleep(0.02)  # Ensure calculation completes
        except:
            try:
                self.app.calculation = 'automatic'
                self.app.calculation = 'manual'
            except:
                self.sheet.calculate()
        
        verified_a = self.sheet.range(output_cell_1).value
        verified_b = self.sheet.range(output_cell_2).value
        
        print(f"\n GLOBAL OPTIMUM FOUND:")
        print(f"   Best region: Region with center ({global_best_region.center[0]*10:.0f}, {global_best_region.center[1]*10:.0f})")
        print(f"   Optimal solution: X1={verification_x1}, X2={verification_x2}")
        print(f"    CACHED values: A={global_best_region.best_a_value:.6f}, B={global_best_region.best_b_value:.6f}")
        print(f"    VERIFIED values: A={verified_a:.6f}, B={verified_b:.6f}")
        print(f"    Constraint check: A ∈ [25.000, 25.100] = {25.000 <= verified_a <= 25.100}")
        print(f"    Input verification: X1={verification_x1} (multiple of 10: {int(verification_x1) % 10 == 0})")
        print(f"    Input verification: X2={verification_x2} (multiple of 10: {int(verification_x2) % 10 == 0})")
        
        # Print all regions for comparison
        print(f"\n ALL FEASIBLE REGIONS FOUND:")
        for i, region in enumerate(sorted(self.feasible_regions, key=lambda r: r.best_b_value)):
            status = " GLOBAL" if region == global_best_region else " LOCAL"
            print(f"   Region {i+1}: {status} | B={region.best_b_value:.4f} | Solution=({region.best_solution[0]*10:.0f},{region.best_solution[1]*10:.0f})")
        
        # Use verified values for final result
        return OptimizationResult(
            verification_x1,              # x1 (verified)
            verification_x2,              # x2 (verified)
            verified_a,                   # A (verified)
            verified_b,                   # B (verified)
            self.eval_count,              # Total evaluations
            len(self.feasible_regions),   # Regions found
            True,                         # Success
            "Single-Strategy DE Optimizer"  # Algorithm name
        )

# =============================================================================
# MAIN EXECUTION
# =============================================================================
if __name__ == "__main__":
    print(" INITIALIZING SINGLE-STRATEGY OPTIMIZER WITH ENHANCED DE")
    
    optimizer = SingleStrategyFeasibilityOptimizer(
        feasibility_tolerance=0.1,      # A ∈ [25.000, 25.100]
        min_region_distance=20          # Minimum distance between regions
    )
    
    try:
        overall_start = time.time()   #  Start global runtime timer
        result = optimizer.optimize(max_evaluations=5000)
        overall_end = time.time()
        
        if result.success:
            print(f"\n SUCCESS: Found optimal solution across {result.feasible_regions_found} feasible regions")
            print(f" Final solution inputs are multiples of 10: X1={result.best_x1}, X2={result.best_x2}")
            print(f" Optimal outputs: A={result.best_a:.6f}, B={result.best_b:.6f}")
            
            print(f"\n MANUAL VERIFICATION RECOMMENDED:")
            print(f"   Please manually input X1={result.best_x1} and X2={result.best_x2} into Excel")
            print(f"   Expected outputs: A≈{result.best_a:.4f}, B≈{result.best_b:.4f}")
        else:
            print(f"\n FAILED: {result.algorithm_used}")
        
        #  Add global runtime print here
        print(f"\n TOTAL RUNTIME: {overall_end - overall_start:.1f} seconds")
            
    except Exception as e:
        print(f"\n CRITICAL ERROR: {str(e)}")
        print(" Try the following troubleshooting steps:")
        print("   1. Ensure Excel file 'v1.xlsx' exists in the current directory")
        print("   2. Ensure the 'Fertigation' sheet exists")
        print("   3. Close any open Excel instances")
        print("   4. Check that cells L4, L5, D26, O11 exist and are accessible")