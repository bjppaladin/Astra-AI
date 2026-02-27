import {
  Shield,
  BarChart3,
  Zap,
  ArrowRight,
  CheckCircle2,
  Users,
  FileText,
  Lock,
} from "lucide-react";
import { Button } from "@/components/ui/button";

export default function Landing() {
  return (
    <div className="min-h-screen bg-background flex flex-col font-sans text-foreground">
      <header className="sticky top-0 z-10 bg-card/80 backdrop-blur-md border-b border-border px-6 py-4">
        <div className="max-w-6xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="h-8 w-8 rounded-md bg-primary flex items-center justify-center text-primary-foreground font-bold">
              A
            </div>
            <span className="text-xl font-semibold tracking-tight font-display">Astra</span>
          </div>
          <a href="/api/login">
            <Button size="sm" className="gap-2" data-testid="button-header-login">
              Sign In
              <ArrowRight className="h-4 w-4" />
            </Button>
          </a>
        </div>
      </header>

      <main className="flex-1">
        <section className="py-20 md:py-32 px-6">
          <div className="max-w-6xl mx-auto grid lg:grid-cols-2 gap-12 items-center">
            <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700">
              <div className="space-y-4">
                <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-primary/10 text-primary text-sm font-medium">
                  <Zap className="h-3.5 w-3.5" />
                  AI-Powered License Intelligence
                </div>
                <h1 className="text-4xl md:text-5xl lg:text-6xl font-display font-bold tracking-tight leading-[1.1]">
                  Optimize your <span className="text-primary">Microsoft 365</span> licenses with confidence
                </h1>
                <p className="text-lg text-muted-foreground max-w-lg leading-relaxed">
                  Astra analyzes your tenant's license assignments and usage patterns, then delivers CIO-level recommendations to cut costs and strengthen security.
                </p>
              </div>

              <div className="flex flex-col sm:flex-row gap-3">
                <a href="/api/login">
                  <Button size="lg" className="gap-2 text-base px-6" data-testid="button-hero-login">
                    Get Started
                    <ArrowRight className="h-4 w-4" />
                  </Button>
                </a>
              </div>

              <div className="flex items-center gap-6 text-sm text-muted-foreground">
                <div className="flex items-center gap-1.5">
                  <CheckCircle2 className="h-4 w-4 text-green-500" />
                  Free to use
                </div>
                <div className="flex items-center gap-1.5">
                  <Lock className="h-4 w-4 text-green-500" />
                  Read-only access
                </div>
                <div className="flex items-center gap-1.5">
                  <Shield className="h-4 w-4 text-green-500" />
                  Data stays private
                </div>
              </div>
            </div>

            <div className="hidden lg:block animate-in fade-in slide-in-from-right-8 duration-700 delay-200">
              <div className="relative rounded-2xl border border-border/60 bg-card p-6 shadow-xl">
                <div className="space-y-4">
                  <div className="flex items-center gap-3 mb-6">
                    <div className="h-3 w-3 rounded-full bg-red-400" />
                    <div className="h-3 w-3 rounded-full bg-yellow-400" />
                    <div className="h-3 w-3 rounded-full bg-green-400" />
                    <div className="ml-auto text-xs text-muted-foreground font-mono">Astra Dashboard</div>
                  </div>
                  <div className="grid grid-cols-3 gap-3">
                    <div className="rounded-lg bg-muted/50 p-3 space-y-1">
                      <div className="text-xs text-muted-foreground">Active Users</div>
                      <div className="text-xl font-bold font-display">247</div>
                    </div>
                    <div className="rounded-lg bg-muted/50 p-3 space-y-1">
                      <div className="text-xs text-muted-foreground">Storage Used</div>
                      <div className="text-xl font-bold font-display">1.8 TB</div>
                    </div>
                    <div className="rounded-lg bg-primary/10 p-3 space-y-1">
                      <div className="text-xs text-muted-foreground">Monthly Cost</div>
                      <div className="text-xl font-bold font-display text-primary">$8,420</div>
                    </div>
                  </div>
                  <div className="space-y-2">
                    <div className="flex justify-between text-xs">
                      <span className="text-muted-foreground">Potential savings (Security strategy)</span>
                      <span className="text-green-500 font-medium">-$1,240/mo</span>
                    </div>
                    <div className="h-2 rounded-full bg-muted overflow-hidden">
                      <div className="h-full rounded-full bg-gradient-to-r from-primary to-primary/60 w-[72%]" />
                    </div>
                  </div>
                  <div className="space-y-1.5 pt-2">
                    {[
                      { label: "Upgrade 12 users to E5 (Security depts)", color: "bg-blue-500" },
                      { label: "Downgrade 8 underutilized E5 licenses", color: "bg-amber-500" },
                      { label: "Remove 15 redundant add-ons", color: "bg-green-500" },
                    ].map((item, i) => (
                      <div key={i} className="flex items-center gap-2 text-xs text-muted-foreground">
                        <div className={`h-1.5 w-1.5 rounded-full ${item.color}`} />
                        {item.label}
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </section>

        <section className="py-16 px-6 border-t border-border/40">
          <div className="max-w-6xl mx-auto">
            <div className="text-center mb-12">
              <h2 className="text-2xl md:text-3xl font-display font-bold mb-3">Everything you need for license optimization</h2>
              <p className="text-muted-foreground max-w-xl mx-auto">
                Connect your Microsoft 365 tenant or upload exports. Astra handles the rest.
              </p>
            </div>
            <div className="grid md:grid-cols-3 gap-6">
              {[
                {
                  icon: BarChart3,
                  title: "Usage-Aware Analysis",
                  desc: "Merges user assignments with mailbox usage data for recommendations based on actual behavior, not just license names.",
                },
                {
                  icon: Users,
                  title: "Per-User Recommendations",
                  desc: "Four optimization strategies (Security, Cost, Balanced, Custom) with per-user license change recommendations and cost impact.",
                },
                {
                  icon: FileText,
                  title: "Executive Briefings",
                  desc: "AI-generated board-ready reports with risk assessments, implementation roadmaps, and financial projections. Export to PDF.",
                },
              ].map((feature, i) => (
                <div
                  key={i}
                  className="p-6 rounded-xl border border-border/60 bg-card hover:border-primary/30 hover:bg-primary/5 transition-all duration-300 group"
                  style={{ animationDelay: `${i * 100}ms` }}
                >
                  <div className="w-10 h-10 rounded-lg bg-primary/10 flex items-center justify-center mb-4 group-hover:bg-primary/20 transition-colors">
                    <feature.icon className="h-5 w-5 text-primary" />
                  </div>
                  <h3 className="font-semibold font-display text-lg mb-2">{feature.title}</h3>
                  <p className="text-sm text-muted-foreground leading-relaxed">{feature.desc}</p>
                </div>
              ))}
            </div>
          </div>
        </section>
      </main>

      <footer className="py-6 text-center text-xs text-muted-foreground border-t border-border/30 px-6">
        <div className="max-w-6xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-2">
          <div>&copy; 2026 Cavaridge, LLC. All rights reserved.</div>
          <div className="flex items-center gap-1 text-muted-foreground/60">
            <span>Powered by</span>
            <span className="font-medium">Microsoft Graph API</span>
          </div>
        </div>
      </footer>
    </div>
  );
}
