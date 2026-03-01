# CPG Annual Sales & Commodity Dashboard

An interactive data visualization dashboard for analyzing Consumer Packaged Goods (CPG) supplier sales performance and commodity cost trends.

## Features

- **Supplier Sales Analysis**: View annual sales data by supplier with interactive charts
- **Commodity Index Tracking**: Monitor commodity price indices over time
- **Cost Impact Analysis**: Analyze how commodity price changes impact overall costs
- **Unit Cost Trends**: Track supplier unit cost trends across multiple dimensions
- **Interactive Charts**: Hover, zoom, and filter data in real-time using Chart.js

## Live Demo

📊 **[View Dashboard Online](https://dseant-create.github.io/PracticumIIdashboard/dashboard.html)**

## Getting Started

### Option 1: View Online (Easiest)
Simply visit the link above - no installation required!

### Option 2: Run Locally

#### Prerequisites
- Python 3.x installed
- Git (optional, for cloning the repository)

#### Installation

1. **Clone or download the repository**
   ```bash
   git clone https://github.com/dseant-create/PracticumIIdashboard.git
   cd PracticumIIdashboard
   ```

2. **Start a local web server**
   ```bash
   python3 -m http.server 8001
   ```
   
   Or use the convenience script (macOS/Linux):
   ```bash
   ./dashboard
   ```

3. **Open in browser**
   ```
   http://localhost:8001/dashboard.html
   ```

## Project Structure

```
├── dashboard.html                    # Main dashboard application
├── supplier_sales_data.json         # Supplier sales figures by year
├── commodity_index_data.json        # Commodity price index trends
├── commodity_item_count.json        # Item count by commodity type
├── commodity_cost_impact.json       # Cost impact analysis data
├── supplier_unit_cost_trends.json   # Unit cost trend data
├── CPG Data Generator.py            # Data generation scripts
├── CPG data Generator_V2_1.25.26.py
├── generate_commodity_index_pack.py
├── generate_supplier_data.py
└── cpg_commodity_pack_out/          # Commodity analysis outputs
    ├── commodity_master.csv
    ├── item_commodity_bridge.csv
    └── other analysis files
```

## Data Files

All data is stored in JSON format for easy integration:

- **supplier_sales_data.json**: Annual sales by supplier (2023-2025)
- **commodity_index_data.json**: Monthly commodity price indices
- **commodity_item_count.json**: Number of items per commodity category
- **commodity_cost_impact.json**: Cost impact metrics
- **supplier_unit_cost_trends.json**: Unit cost changes by supplier

## Usage

1. **View Supplier Sales**: The main dashboard displays top suppliers by 2025 sales
2. **Analyze Trends**: Scroll down to see commodity index trends and cost impacts
3. **Interactive Charts**: 
   - Hover over data points for details
   - Click legend items to toggle visibility
   - Use zoom and pan tools for closer inspection

## Technologies Used

- **Frontend**: HTML5, CSS3, JavaScript (ES6+)
- **Charting**: [Chart.js](https://www.chartjs.org/) with Data Labels plugin
- **Data Format**: JSON
- **Hosting**: GitHub Pages

## Development

### Local Development Setup

1. Install dependencies (if using Node.js for tooling):
   ```bash
   npm install
   ```

2. Make changes to `dashboard.html` or data JSON files

3. Test locally by running the web server (see Getting Started)

4. Push changes to GitHub:
   ```bash
   git add .
   git commit -m "Description of changes"
   git push origin main
   ```

Changes will automatically deploy to GitHub Pages within a few minutes.

### Generating New Data

To regenerate data files:
```bash
python3 "CPG Data Generator.py"
```

## Troubleshooting

### Dashboard not loading?
- Ensure you're using `http://` (not `https://`) for local development
- Clear browser cache (Cmd+Shift+R on macOS)
- Check browser console (F12) for detailed error messages
- Verify all JSON files are in the same directory as dashboard.html

### Data not displaying?
- Verify JSON files exist and are valid
- Check browser console for fetch errors
- Ensure web server is running (you should see server logs)

## License

This project is part of the MABA Practicum II coursework.

## Contact

For questions or feedback, contact the project maintainer.

---

**Last Updated**: February 6, 2026
