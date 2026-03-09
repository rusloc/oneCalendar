---
name: PBI-Custom-Visual-Creator-Skill
description: Used when asked to create a special custom visual for Power BI. Use when asked to develop or amend / fix custom power BI visual using pbiviz tool.
---

# Power BI Custom Visual Creator Skill

When asked to create a custom visual for Power BI use these steps:

## Steps:

1.	Setup environment. If needed ask user for permissions and/or ask user to run commands
2.	Learn structure of the default/template project: https://learn.microsoft.com/en-us/power-bi/developer/visuals/visual-project-structure
3.	Setup default project structure. If needed ask user to run commands or provide permissions
4.	After setup/init check if the template project structure is full and correct
5.	Learn Core visual logic / code: https://learn.microsoft.com/en-us/power-bi/developer/visuals/visual-api
6.	Check newest visual API version: https://github.com/microsoft/powerbi-visuals-api/blob/main/src/visuals-api.d.ts
7.	Learn Core visual settings (properties): https://learn.microsoft.com/en-us/power-bi/developer/visuals/objects-properties
8.	Learn how to add formatting options: https://learn.microsoft.com/en-us/power-bi/developer/visuals/custom-visual-develop-tutorial-format-options
9.	Learn about Capabilities: https://learn.microsoft.com/en-us/power-bi/developer/visuals/capabilities
10.	Learn how to create and edit dataViewMappings: https://learn.microsoft.com/en-us/power-bi/developer/visuals/dataview-mappings
11.	Learn how to create a simple KPI card custom visual: https://learn.microsoft.com/en-us/power-bi/developer/visuals/develop-circle-card
12.	Learn how to create simple bar chart: https://learn.microsoft.com/en-us/power-bi/developer/visuals/create-bar-chart?tabs=CreateNewVisual
13.	Create the Core visual logic / code according to the user definition/task. Core template code in ./src/visual.ts
14.	Create Core visual settings (properties) to match the user's definition/task. Core template code in ./src/settings.ts
15.	Define basic capabilities to match user's request
16.	Define dataViewMappings if needed 
17.	Setup basic formatting options. Core formatting options in ./style/visual.less

# References

Always check reference.md when creating/amending custom visual. Before completing any user request check if any link from reference.md is applicable to the task.
If link is applicable: read, learn and apply knowledge.

