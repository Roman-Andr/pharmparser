import { useQuery } from "@tanstack/react-query";
import { api } from "../api";
import type { Bootstrap, Run } from "../types";
export const useBootstrap = () => useQuery({ queryKey: ["bootstrap"], queryFn: () => api<Bootstrap>("/bootstrap") });
export const useHistory = () => useQuery({ queryKey: ["history"], queryFn: () => api<Run[]>("/history"), refetchInterval: 2000 });
