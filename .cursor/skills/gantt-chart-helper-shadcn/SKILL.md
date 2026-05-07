---
name: gantt-chart-helper-shadcn
description: >-
  The gantt-chart-helper repository uses shadcn/ui for app UI components. Use
  when building or refactoring UI, forms, dialogs, tables, tabs, and buttons.
  Prefer shadcn/ui components and styling patterns over custom pure CSS unless
  the user explicitly requests custom styling.
---

# gantt-chart-helper: use shadcn/ui

## Component preference

- Prefer components in `src/components/ui/*` and shadcn/ui patterns for new UI work.
- For refactors, replace custom controls with shadcn/ui equivalents where practical.
- Keep behavior in React components and use utility classes for styling consistency.

## Typical mapping

- Buttons -> `Button`
- Text/date/number inputs -> `Input`
- Tabbed UI -> `Tabs`, `TabsList`, `TabsTrigger`, `TabsContent`
- Dense data layout -> `Table` family components
- Grouped sections -> `Card`, `CardHeader`, `CardContent`, `CardFooter`

## Do not

- Introduce or expand large one-off pure CSS systems when a shadcn/ui component can cover the same UI.
- Create duplicate custom base components that overlap with shadcn/ui primitives.
