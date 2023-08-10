""" -----------------------------------------------------------------------------------------------------
This is the python script to calculate multiple variables in ANSYS Electronics to reduce the time spend on process
** Scripting.pdf section 21 may help **
first line imports the script environment
second line initializes the script environment on ANSYS Electronics Desktop
third line restores the ANSYS Electronics Desktop window if it was minimized or not visible.
fourth and fifth lines set the active project and active design within the project. The project is named
"c_core_for_matlab_modified", and the design is named "c_core_for_matlab_modified_for_RAC".
sixth line gets access to the "FieldsReporter" module within the active design. This module is likely used
 for post-processing and reporting on electromagnetic field results.
----------------------------------------------------------------------------------------------------------"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()
oProject = oDesktop.SetActiveProject("c_core_for_matlab_modified")
oDesign = oProject.SetActiveDesign("c_core_for_matlab_modified_for_RAC")
oModule = oDesign.GetModule("FieldsReporter")

num_of_circles = 12  # number of circle surfaces that the resistance will be calculated
length = 25 * (10 ** -2)  # copper cable length in one turn (meter)
rho = 1.724 * (10 ** -8)  # copper resistivity (ohm*meter)
current = 1.5  # current flows through the conductors
N = length / pow(current, 2)  # comes from the formula
initial_name = "Circle"  # base name of the surface in ANSYS

for i in range(1, num_of_circles+1):
    surface_name = initial_name + str(i)
    oModule.CopyNamedExprToStack("J_Vector")
    oModule.CopyNamedExprToStack("J_Vector")
    oModule.CalcOp("Dot")  # Dot product operation
    oModule.EnterScalar(rho)
    oModule.CalcOp("*")
    oModule.EnterSurf(surface_name)
    oModule.CalcOp("Integrate")
    oModule.EnterScalar(N)
    oModule.CalcOp("*")  # multiplying operation
    oModule.AddNamedExpression("AC_resistance_" + str(i), "Fields")  # AC_resistance_1, AC_resistance_2 etc
    oModule.CopyNamedExprToStack("AC_resistance_" + str(i))  # copy to stack button
    oModule.ClcEval("Setup1 : LastAdaptive",
        [
            "Freq:="		, "10kHz",
            "I:="			, "1.5A",
            "N:="			, "100",
            "Phase:="		, "0deg",
            "g:="			, "3mm"
        ], "Fields")  # evaluate the result for the parameters
""" -----------------------------------------------------------------------------------------------------
Formula for AC resistance calculation is:
R_ac = [length*rho/(current)^2]*(surface_integral(J^2.dS))
----------------------------------------------------------------------------------------------------------"""