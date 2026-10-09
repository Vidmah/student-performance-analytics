# Intelligent Student Performance Analytics

React + TypeScript + Tailwind CSS dashboard combining Apriori association rules, J48 decision tree,
Naive Bayes, linear regression and K-Means clustering. Runs entirely in the browser.

## Run it
```
npm install
npm run dev        # http://localhost:5173
npm run build      # production build in dist/
```

## Quick option
Open `standalone/student-analytics.html` in a browser (needs internet for the CDN scripts). No install required.

## Structure
- `src/App.tsx`   – algorithms (Apriori, J48, Naive Bayes, regression, K-Means), charts and all pages
- `src/index.css` – Tailwind layers plus the glassmorphism theme
- `src/main.tsx`  – entry point

## Using your own data
Upload a CSV/Excel file on the Dataset page. It needs a header row and a numeric marks column (e.g. `Final_Marks`).
