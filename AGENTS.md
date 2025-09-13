# Agent Guidelines for c-toolkit

## Build Commands

- `bun run build` - Build for production
- `bun run lint` - Run ESLint
- No test framework configured
- **IMPORTANT**: Never run `bun run dev`, `bun run build` or any start commands

## Tech Stack

- Next.js 15 + React 19 + TypeScript
- Convex backend with Clerk authentication
- Tailwind CSS + Prettier formatting + shadcn/ui components
- Bun package manager

## UI Components (shadcn/ui)

- **Installation**: Use `npx shadcn@latest add [component]` to add components
- **Component Structure**: Located in `components/ui/` directory
- **Customization**: Components are meant to be copied and customized, not imported from npm
- **Styling**: Built with Tailwind CSS and Radix UI primitives
- **Accessibility**: All components follow WCAG guidelines and use proper ARIA attributes
- **Theming**: Supports light/dark mode via CSS variables in `globals.css`
- **Available Components**:
  - Layout: `sidebar`, `sheet`, `separator`, `collapsible`
  - Navigation: `breadcrumb`, `dropdown-menu`
  - Form: `button`, `input`
  - Feedback: `tooltip`, `skeleton`, `avatar`

## Code Style

- Follow Next.js/ESLint config conventions
- Use Prettier defaults (empty .prettierrc)
- Strict TypeScript with ES2017 target
- Import aliases: `@/*` maps to root
- **shadcn/ui**: Import components from `@/components/ui/[component]`
- **Component Patterns**: Use composition over configuration for UI components
- Convex functions: always include `args` and `returns` validators using `v.*`
- Use new Convex syntax with explicit function definitions
- File organization: components/, app/, convex/ structure

## UI Development Guidelines

- **Component Usage**: Always prefer shadcn/ui components over custom implementations
- **Accessibility**: Ensure all interactive elements have proper focus states and ARIA labels
- **Responsive Design**: Use Tailwind's responsive utilities (`sm:`, `md:`, `lg:`, `xl:`)
- **Theming**: Use CSS custom properties for colors, defined in `globals.css`
- **Icons**: Use Lucide React icons (default with shadcn/ui)
- **Component Variants**: Use `cva` (class-variance-authority) for component variants
- **Animation**: Prefer CSS transitions and Framer Motion when needed

## Convex Guidelines

- Follow extensive rules in `.cursor/rules/convex_rules.mdc`
- Use `query`, `mutation`, `action` for public APIs
- Use `internalQuery`, `internalMutation`, `internalAction` for private functions
- Always validate arguments with `v.*` validators
- Import from `./_generated/server` and `./_generated/api`
- Use proper function references for cross-function calls
