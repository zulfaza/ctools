# CTools

<img src="public/logo.png" alt="CTools" width="270" height="96">

Your Internal Tools, in One Place. A simple hub to access dashboards and utilities that keeps tools discoverable and easy to reach.

## Features

- **Excel Reports Processor** - Upload one or more spreadsheets to generate combined metrics and summaries
- **Collective Subscription Management** - Manage shared subscriptions and split costs among team members _(coming soon)_
- **Main Dashboard** - Overview and quick access to currently available utilities
- **Misc Utilities** - A growing set of helper tools with functions that vary by use-case

## Tech Stack

- [Next.js 15](https://nextjs.org/) with React 19 for the frontend
- [Convex](https://convex.dev/) as the backend (database, server logic)
- [Clerk](https://clerk.com/) for authentication
- [Tailwind CSS](https://tailwindcss.com/) with [shadcn/ui](https://ui.shadcn.com/) components
- [Bun](https://bun.sh/) as the package manager

## Get started

If you just cloned this codebase and didn't use `npm create convex`, run:

```
bun install
bun run dev
```

If you're reading this README on GitHub and want to use this template, run:

```
npm create convex@latest -- -t nextjs-clerk
```

Then:

1. Open your app. There should be a "Claim your application" button from Clerk in the bottom right of your app.
2. Follow the steps to claim your application and link it to this app.
3. Follow step 3 in the [Convex Clerk onboarding guide](https://docs.convex.dev/auth/clerk#get-started) to create a Convex JWT template.
4. Uncomment the Clerk provider in `convex/auth.config.ts`
5. Paste the Issuer URL as `CLERK_JWT_ISSUER_DOMAIN` to your dev deployment environment variable settings on the Convex dashboard (see [docs](https://docs.convex.dev/auth/clerk#configuring-dev-and-prod-instances))

If you want to sync Clerk user data via webhooks, check out this [example repo](https://github.com/thomasballinger/convex-clerk-users-table/).

## Learn more

To learn more about developing your project with Convex, check out:

- The [Tour of Convex](https://docs.convex.dev/get-started) for a thorough introduction to Convex principles.
- The rest of [Convex docs](https://docs.convex.dev/) to learn about all Convex features.
- [Stack](https://stack.convex.dev/) for in-depth articles on advanced topics.

## Join the community

Join thousands of developers building full-stack apps with Convex:

- Join the [Convex Discord community](https://convex.dev/community) to get help in real-time.
- Follow [Convex on GitHub](https://github.com/get-convex/), star and contribute to the open-source implementation of Convex.
