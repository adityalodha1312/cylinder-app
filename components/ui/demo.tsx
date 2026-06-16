"use client"

import { useState } from "react"
import { LayoutGrid, ClipboardList, BarChart3, Search, Users } from "lucide-react"
import { MenuBar } from "@/components/ui/glow-menu"

const menuItems = [
  {
    icon: LayoutGrid,
    label: "Home",
    href: "/admin/dashboard",
    gradient:
      "radial-gradient(circle, rgba(59,130,246,0.15) 0%, rgba(37,99,235,0.06) 50%, rgba(29,78,216,0) 100%)",
    iconColor: "text-blue-500",
  },
  {
    icon: ClipboardList,
    label: "Activity",
    href: "/admin/activity",
    gradient:
      "radial-gradient(circle, rgba(249,115,22,0.15) 0%, rgba(234,88,12,0.06) 50%, rgba(194,65,12,0) 100%)",
    iconColor: "text-orange-500",
  },
  {
    icon: BarChart3,
    label: "Inventory",
    href: "/admin/inventory",
    gradient:
      "radial-gradient(circle, rgba(34,197,94,0.15) 0%, rgba(22,163,74,0.06) 50%, rgba(21,128,61,0) 100%)",
    iconColor: "text-green-500",
  },
  {
    icon: Search,
    label: "Search",
    href: "/admin/search",
    gradient:
      "radial-gradient(circle, rgba(139,92,246,0.15) 0%, rgba(139,92,246,0.06) 50%, rgba(139,92,246,0) 100%)",
    iconColor: "text-purple-500",
  },
  {
    icon: Users,
    label: "Customers",
    href: "/admin/customers",
    gradient:
      "radial-gradient(circle, rgba(239,68,68,0.15) 0%, rgba(220,38,38,0.06) 50%, rgba(185,28,28,0) 100%)",
    iconColor: "text-red-500",
  },
]

export function MenuBarDemo() {
  const [activeItem, setActiveItem] = useState<string>("Home")

  return (
    <div className="flex items-center justify-center p-4">
      <MenuBar
        items={menuItems}
        activeItem={activeItem}
        onItemClick={setActiveItem}
      />
    </div>
  )
}
