# Using Excel MCP Server with AI Assistants

This guide provides examples of how to use the Excel MCP Server with different AI assistants.

## Overview

The Excel MCP Server enables AI assistants to:
1. Create, read, and modify Excel files
2. Analyze data and create visualizations
3. Apply formatting and styling to workbooks
4. Generate reports and extract insights

## Integration Examples

### GitHub Copilot

GitHub Copilot can help you write code that interacts with Excel files through the Excel MCP server:

```python
# Example: Analyzing sales data with Copilot
import pandas as pd
import matplotlib.pyplot as plt

# Load data from an Excel file
df = pd.read_excel("sales_data.xlsx")

# Calculate monthly sales trends
monthly_sales = df.groupby(pd.Grouper(key='Date', freq='M'))['Amount'].sum()

# Create a visualization
plt.figure(figsize=(12, 6))
monthly_sales.plot(kind='bar')
plt.title('Monthly Sales Trends')
plt.xlabel('Month')
plt.ylabel('Sales Amount')

# Save the chart to a new Excel file
with pd.ExcelWriter('sales_report.xlsx') as writer:
    df.to_excel(writer, sheet_name='Raw Data')
    monthly_sales.to_excel(writer, sheet_name='Monthly Summary')
    
    # You can ask Copilot to help you add an Excel chart here
```

### Claude and Anthropic API

Claude can analyze Excel data and provide insights when you integrate it with the Excel MCP Server:

```python
import anthropic
import json
import os

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

# Example: Asking Claude to analyze a financial spreadsheet
def analyze_financial_data(filename):
    message = client.messages.create(
        model="claude-3-opus-20240229",
        max_tokens=1000,
        messages=[
            {
                "role": "user",
                "content": f"I have an Excel file named {filename} containing financial data. " 
                           f"The first worksheet has quarterly sales figures. "
                           f"Could you analyze this data and tell me which quarter performed best? "
                           f"Also, create a pivot table showing regional performance."
            }
        ],
    )
    
    return message.content

analysis = analyze_financial_data("financial_report_2023.xlsx")
print(analysis)
```

### OpenAI API with Excel MCP Server

OpenAI's GPT models can be used to analyze and manipulate Excel data through the Excel MCP Server:

```python
import openai
import os

client = openai.OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# Example: Using GPT-4 to create an Excel budget template
def create_budget_template():
    tools = [
        {
            "type": "function",
            "function": {
                "name": "create_excel_workbook",
                "description": "Creates a new Excel workbook",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "filename": {
                            "type": "string",
                            "description": "Name of the Excel file to create"
                        }
                    },
                    "required": ["filename"]
                }
            }
        },
        {
            "type": "function",
            "function": {
                "name": "add_worksheet",
                "description": "Adds a worksheet to an Excel workbook",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "filename": {"type": "string"},
                        "sheet_name": {"type": "string"}
                    },
                    "required": ["filename", "sheet_name"]
                }
            }
        }
    ]
    
    response = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role": "system", "content": "You are a financial expert who creates Excel templates."},
            {"role": "user", "content": "Create a monthly budget template in Excel with income and expense tracking."}
        ],
        tools=tools,
        tool_choice="auto"
    )
    
    return response.choices[0].message

template_creation = create_budget_template()
print(template_creation)
```

## Example Workflows

### Data Analysis Workflow

1. **Upload data**:
   - Upload data through the web interface at http://localhost:8080
   - Upload programmatically via API

2. **Process with AI**:
   - Connect your AI assistant to the MCP server
   - Ask it to analyze the Excel data
   - Request visualizations or trend analysis

3. **Export results**:
   - Have the AI save the results to a new Excel file
   - Download the file through the web interface

### Report Generation

```python
# Example: Generating a monthly sales report with AI
prompt = """
Using the Excel file 'sales_data.xlsx', please:
1. Analyze monthly sales trends
2. Identify top-performing products
3. Create a professional report with charts
4. Save the report as 'monthly_sales_report.xlsx'
"""

# Your AI tool would process this prompt using the Excel MCP server
```

## Tips for Effective AI-Excel Integration

1. **Be specific about file paths**:
   - Always specify the exact filename in the Excel files directory
   - Use forward slashes in paths, even on Windows

2. **Provide context about data structure**:
   - Tell the AI about column names and data types
   - Explain the meaning of different worksheets

3. **Request specific Excel functions**:
   - Ask for VLOOKUP, SUMIF, or other specific Excel functions
   - Specify formatting requirements (currency, dates, etc.)

4. **Leverage visualization capabilities**:
   - Request specific chart types (bar, line, scatter, etc.)
   - Specify labels, colors, and other visualization parameters

## Next Steps

- Check the [TOOLS.md](TOOLS.md) document for a complete list of Excel operations supported
- Join our [Discord community](https://discord.gg/example) for more examples and support
- Contribute your own examples by submitting a pull request
