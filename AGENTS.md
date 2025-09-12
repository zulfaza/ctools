# Agent Guidelines for c-toolkit

## Build Commands

- `bun run dev` - Start development server (Next.js + Convex)
- `bun run build` - Build for production
- `bun run lint` - Run ESLint
- No test framework configured

## Tech Stack

- Next.js 15 + React 19 + TypeScript
- Convex backend with Clerk authentication
- Tailwind CSS + Prettier formatting
- Bun package manager

## Code Style

- Follow Next.js/ESLint config conventions
- Use Prettier defaults (empty .prettierrc)
- Strict TypeScript with ES2017 target
- Import aliases: `@/*` maps to root
- Convex functions: always include `args` and `returns` validators using `v.*`
- Use new Convex syntax with explicit function definitions
- File organization: components/, app/, convex/ structure

## Convex Guidelines

- Follow extensive rules in `.cursor/rules/convex_rules.mdc`
- Use `query`, `mutation`, `action` for public APIs
- Use `internalQuery`, `internalMutation`, `internalAction` for private functions
- Always validate arguments with `v.*` validators
- Import from `./_generated/server` and `./_generated/api`
- Use proper function references for cross-function calls
