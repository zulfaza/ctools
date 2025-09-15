import Link from "next/link";
import { UserButton, SignInButton, SignedIn, SignedOut } from "@clerk/nextjs";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import { FileSpreadsheet, LayoutDashboard, Wrench } from "lucide-react";
import { Metadata } from "next";

export const metadata: Metadata = {
  title: "CTools",
  description:
    "A simple hub to access dashboards and utilities. Tools may be independent — this page keeps them discoverable and easy to reach.",
};

export default function Home() {
  return (
    <>
      <header className="sticky top-0 z-10 bg-background/80 backdrop-blur supports-[backdrop-filter]:bg-background/60 border-b border-slate-200 dark:border-slate-800">
        <div className="mx-auto max-w-6xl px-4 h-14 flex items-center justify-between">
          <Link href="/" className="flex items-center gap-2 font-semibold">
            <span
              className="inline-block h-2 w-2 rounded-sm bg-primary"
              aria-hidden
            />
            <span>ctools</span>
          </Link>
          <nav className="flex items-center gap-2">
            <SignedOut>
              <SignInButton mode="modal">
                <Button variant="outline" size="sm" aria-label="Sign in">
                  Sign in
                </Button>
              </SignInButton>
            </SignedOut>
            <SignedIn>
              <Link href="/dashboard">
                <Button size="sm" aria-label="Go to dashboard">
                  Dashboard
                </Button>
              </Link>
              <UserButton />
            </SignedIn>
          </nav>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-4 py-16">
        {/* Hero */}
        <section className="text-center">
          <h1 className="text-4xl md:text-5xl font-bold tracking-tight">
            Your Internal Tools, in One Place
          </h1>
          <p className="mt-4 text-muted-foreground max-w-2xl mx-auto">
            A simple hub to access dashboards and utilities. Tools may be
            independent — this page keeps them discoverable and easy to reach.
          </p>
          <div className="mt-8 flex items-center justify-center gap-3">
            <Link href="/dashboard">
              <Button size="lg">Open Dashboard</Button>
            </Link>
            <SignedOut>
              <SignInButton mode="modal">
                <Button variant="outline" size="lg">
                  Sign in
                </Button>
              </SignInButton>
            </SignedOut>
          </div>
        </section>

        <Separator className="my-12" />

        {/* Tools grid */}
        <section
          aria-label="Available tools"
          className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3"
        >
          <ToolCard
            href="/dashboard"
            title="Excel Reports"
            description="Upload one or more spreadsheets to generate combined metrics and summaries."
            Icon={FileSpreadsheet}
          />
          <ToolCard
            href="/dashboard"
            title="Main Dashboard"
            description="Overview and quick access to currently available utilities."
            Icon={LayoutDashboard}
          />
          <ToolCard
            href="/dashboard"
            title="Misc Utilities"
            description="A growing set of helper tools. Functions may vary by use-case."
            Icon={Wrench}
          />
        </section>

        {/* Footer */}
        <footer className="mt-16 text-center text-xs text-muted-foreground">
          <p>Built with Next.js, Convex, Clerk, and shadcn/ui.</p>
        </footer>
      </main>
    </>
  );
}

type ToolCardProps = {
  href: string;
  title: string;
  description: string;
  Icon: React.ComponentType<{ className?: string }>;
};

function ToolCard({ href, title, description, Icon }: ToolCardProps) {
  return (
    <Link
      href={href}
      className="group rounded-lg border border-slate-200 dark:border-slate-800 p-5 transition-colors hover:border-primary focus:outline-none focus-visible:ring-2 focus-visible:ring-primary/40"
      aria-label={title}
    >
      <div className="flex items-start gap-4">
        <div className="rounded-md bg-primary/10 text-primary p-2">
          <Icon className="h-5 w-5" />
        </div>
        <div>
          <h3 className="font-medium leading-none mb-1 group-hover:text-primary">
            {title}
          </h3>
          <p className="text-sm text-muted-foreground">{description}</p>
        </div>
      </div>
    </Link>
  );
}
