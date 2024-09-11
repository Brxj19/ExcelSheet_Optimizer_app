#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import pulp as pl

class OptimizationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Optimization Software")
        self.root.geometry("600x400")
        
        # Initialize variables
        self.input_file_path = None
        self.df_variables = None
        self.df_objective = None
        self.df_constraints = None
        self.optimization_result = None
        
        # UI Elements
        self.create_widgets()

    def create_widgets(self):
        # Load Excel file
        load_button = tk.Button(self.root, text="Load Excel File", command=self.load_file, width=20)
        load_button.pack(pady=10)
        
        # Run Optimization
        optimize_button = tk.Button(self.root, text="Run Optimization", command=self.run_optimization, width=20)
        optimize_button.pack(pady=10)
        
        # Save Results
        save_button = tk.Button(self.root, text="Save Results", command=self.save_results, width=20)
        save_button.pack(pady=10)
        
        # Display results
        self.result_label = tk.Label(self.root, text="", wraplength=500)
        self.result_label.pack(pady=20)

    def load_file(self):
        self.input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if self.input_file_path:
            try:
                # Load all necessary sheets
                self.df_variables = pd.read_excel(self.input_file_path, sheet_name='Variables').astype(str)
                self.df_objective = pd.read_excel(self.input_file_path, sheet_name='Objective').astype(str)
                self.df_constraints = pd.read_excel(self.input_file_path, sheet_name='Constraints').astype(str)
                messagebox.showinfo("File Loaded", "Excel file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {e}")
        print(self.df_variables)
        
    def run_optimization(self):
        if self.df_variables is None or self.df_objective is None or self.df_constraints is None:
            messagebox.showerror("Error", "Please load an Excel file with the correct sheets.")
            return
        
        try:
            # Create a problem instance
            prob = pl.LpProblem("Dynamic_Problem", pl.LpMaximize)
            
            # Create decision variables
            variables = {}
            for _, row in self.df_variables.iterrows():
                var_name = row['Variable']
                lb = row['LowerBound']
                ub = row['UpperBound'] if pd.notnull(row['UpperBound']) else None
                cat = pl.LpContinuous if row['Category'] == 'Continuous' else pl.LpInteger
                variables[var_name] = pl.LpVariable(var_name, lowBound=lb, upBound=ub, cat=cat)
            
            # Define the objective function
            prob += pl.lpSum(self.df_objective['Coefficient'][i] * variables[var] 
                             for i, var in enumerate(self.df_objective['Variable']))

            # Add constraints
            for constraint_name in self.df_constraints['Constraint'].unique():
                constraint_df = self.df_constraints[self.df_constraints['Constraint'] == constraint_name]
                prob += (pl.lpSum(constraint_df['Coefficient'].iloc[i] * variables[var] 
                                  for i, var in enumerate(constraint_df['Variable'])) 
                         <= constraint_df['RHS'].iloc[0]), constraint_name
            
            # Solve the problem
            prob.solve()
            
            # Retrieve results
            self.optimization_result = {var: variables[var].varValue for var in variables}
            self.optimization_result['Objective'] = pl.value(prob.objective)
            
            # Display results
            result_text = "Optimization Results:\n\n" + "\n".join([f"{var} = {value}" for var, value in self.optimization_result.items()])
            self.result_label.config(text=result_text)
            messagebox.showinfo("Optimization Complete", "Optimization completed successfully!")
        
        except Exception as e:
            messagebox.showerror("Error", f"Optimization failed: {e}")
        
    def save_results(self):
        if self.optimization_result is None:
            messagebox.showerror("Error", "No results to save. Please run the optimization first.")
            return
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_path:
            try:
                # Convert results to DataFrame
                result_df = pd.DataFrame([self.optimization_result])
                # Save results to Excel
                result_df.to_excel(save_path, index=False)
                messagebox.showinfo("Save Successful", "Results saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OptimizationApp(root)
    root.mainloop()

