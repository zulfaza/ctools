"use client";

import * as React from "react";
import { FileChartColumnIncreasing } from "lucide-react";
import { OrganizationSwitcher } from "@clerk/nextjs";

import { NavMain } from "@/components/nav-main";
import { NavUser } from "@/components/nav-user";
import {
  Sidebar,
  SidebarContent,
  SidebarFooter,
  SidebarHeader,
} from "@/components/ui/sidebar";

const data = {
  navMain: [
    {
      title: "Excel Reports",
      url: "/dashboard",
      icon: FileChartColumnIncreasing,
    },
  ],
};

export function AppSidebar({ ...props }: React.ComponentProps<typeof Sidebar>) {
  return (
    <Sidebar variant="inset" {...props}>
      <SidebarHeader className="space-y-2">
        <div className="flex items-center gap-2 rounded-md px-1.5 py-1.5 ">
          <div
            className="h-10 w-10 rounded-lg bg-sidebar-accent text-sidebar-accent-foreground grid place-items-center"
            aria-hidden
          >
            <FileChartColumnIncreasing className="size-5" />
          </div>
          <span className="text-lg font-semibold tracking-tight">Ctools</span>
        </div>
        {/* Organization switcher (smaller) */}
        <OrganizationSwitcher
          hidePersonal
          afterCreateOrganizationUrl="/dashboard"
          afterSelectOrganizationUrl="/dashboard"
          appearance={{
            variables: {
              colorBackground: "var(--sidebar)",
              colorInputBackground: "var(--sidebar)",
              colorText: "var(--sidebar-foreground)",
              colorTextSecondary: "var(--sidebar-accent-foreground)",
              colorNeutral: "var(--sidebar-accent-foreground)",
              borderRadius: "0.5rem",
              fontSize: "0.875rem",
            },
            elements: {
              rootBox: "w-full! mt-0.5",
              organizationSwitcherTrigger:
                "w-full px-4! py-3! border-2! justify-between! bg-transparent hover:bg-sidebar-accent hover:text-sidebar-accent-foreground rounded-lg data-[state=open]:bg-sidebar-accent data-[state=open]:text-sidebar-accent-foreground",
              organizationSwitcherTriggerIcon: "size-6 rounded-lg",
              organizationPreviewTextContainer: "flex-1 text-left",
              organizationPreviewMainIdentifier: "text-xs font-medium truncate",
              organizationPreviewSecondaryIdentifier: "text-[10px] truncate",
            },
          }}
        />
      </SidebarHeader>
      <SidebarContent>
        <NavMain items={data.navMain} />
      </SidebarContent>
      <SidebarFooter>
        <NavUser />
      </SidebarFooter>
    </Sidebar>
  );
}
