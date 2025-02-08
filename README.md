# Optis Screen Capture Automation
 

**Problem Statement**
The manual capture of Optis green screen (mainframe) screenshots has been a time-consuming daily task for Citi’s testing and remediation teams. This activity is essential for creating pass logs required by the testing team and serves as evidence for audit compliance.

Existing solutions using UFT were available within the team but incurred significant licensing costs, which limited scalability and added operational expenses.

**Proposed Solution**
To bypass the license and user dependency, a solution was considered to automate using Excel VBA.

User needs to input the account numbers, screens that need to be captured, name and path of the Word file into the macro’s input sheet and run it.

This macro connects to the PComm session, signs in to the account, captures required screenshots, and saves them to the Word file in the path given in the macro.

**Solution Benefits**
Cost Efficiency: Eliminates licensing costs by using readily available Excel and VBA, making the solution highly cost-effective.

Accuracy and Reliability: Automation minimizes the risk of human error, ensuring consistency in screenshot captures for each session.

Significant Time Savings: Achieves an 80% reduction in the time required for screenshot capture, allowing the remediation team to focus on higher-value tasks and increasing overall productivity.
