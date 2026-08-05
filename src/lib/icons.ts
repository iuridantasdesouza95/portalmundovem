import {
  BarChart3,
  Bell,
  CalendarRange,
  Factory,
  FileText,
  Gauge,
  LayoutGrid,
  LineChart,
  PieChart,
  Settings,
  Target,
  TrendingUp,
  Truck,
  Users,
  Wallet,
  type LucideIcon,
} from "lucide-react";

export const iconMap: Record<string, LucideIcon> = {
  BarChart3,
  Bell,
  CalendarRange,
  Factory,
  FileText,
  Gauge,
  LayoutGrid,
  LineChart,
  PieChart,
  Settings,
  Target,
  TrendingUp,
  Truck,
  Users,
  Wallet,
};

export const iconNames = Object.keys(iconMap);

export function getIcon(name?: string | null): LucideIcon {
  return (name && iconMap[name]) || BarChart3;
}
