import React from "react";
import ReactDOM from "react-dom/client";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import { createBrowserRouter, Navigate, RouterProvider } from "react-router-dom";
import { AppLayout } from "./ui/AppLayout";
import { HomePage } from "./ui/HomePage";
import { ProfilesPage } from "./ui/ProfilesPage";
import { ExecutionPage } from "./ui/ExecutionPage";
import { HistoryPage } from "./ui/HistoryPage";
import { SettingsPage } from "./ui/SettingsPage";
import { HelpPage } from "./ui/HelpPage";
import { OnboardingPage } from "./ui/OnboardingPage";
import "./styles.css";

const router = createBrowserRouter([{ path: "/", element: <AppLayout />, children: [
  { index: true, element: <HomePage /> }, { path: "profiles", element: <ProfilesPage /> },
  { path: "run/:runId", element: <ExecutionPage /> }, { path: "history", element: <HistoryPage /> },
  { path: "settings", element: <SettingsPage /> }, { path: "help", element: <HelpPage /> },
  { path: "onboarding", element: <OnboardingPage /> }, { path: "*", element: <Navigate to="/" replace /> },
]}]);
const queryClient = new QueryClient({ defaultOptions: { queries: { staleTime: 5000, retry: 1 } } });
ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode><QueryClientProvider client={queryClient}><RouterProvider router={router} /></QueryClientProvider></React.StrictMode>,
);
