# Agency SpendPlan Tool

Browser-based budget execution and spend-plan reporting tool. Users upload the Spend Plan and APE Lookup workbooks locally; workbook data is processed in the browser and is not stored by the application.

## Local development

```bash
pnpm install
pnpm dev
```

## Production build

```bash
pnpm build
```

Deploy the generated `dist` directory with Azure Static Web Apps.
