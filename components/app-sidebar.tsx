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
      <SidebarHeader>
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
              fontSize : "1rem",
            },
            elements: {
              rootBox: "w-full",
              organizationSwitcherTrigger: "w-full justify-start bg-transparent hover:bg-sidebar-accent hover:text-sidebar-accent-foreground p-2 rounded-lg data-[state=open]:bg-sidebar-accent data-[state=open]:text-sidebar-accent-foreground",
              organizationSwitcherTriggerIcon: "size-8 rounded-lg",
              organizationPreviewTextContainer: "flex-1 text-left",
              organizationPreviewMainIdentifier: "text-sm font-medium truncate",
              organizationPreviewSecondaryIdentifier: "text-xs truncate",
              avatarBox : "h-10! w-10!"
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
