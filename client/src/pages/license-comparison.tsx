import { useState, useEffect, useMemo, Fragment } from "react";
import { useLocation, useSearch } from "wouter";
import {
  Check,
  X as XIcon,
  Minus,
  ArrowLeft,
  Trash2,
  Plus,
  ExternalLink,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  LICENSES,
  FEATURE_CATEGORIES,
  type LicenseInfo,
  type FeatureValue,
} from "@/lib/license-data";

const EMPTY = "__empty__";

function renderFeatureValue(value: FeatureValue) {
  if (value === true) {
    return (
      <div className="flex items-center justify-center" data-testid="feature-included">
        <div className="h-6 w-6 rounded-full bg-green-500/10 flex items-center justify-center">
          <Check className="h-4 w-4 text-green-600" />
        </div>
      </div>
    );
  }
  if (value === false) {
    return (
      <div className="flex items-center justify-center" data-testid="feature-not-included">
        <div className="h-6 w-6 rounded-full bg-muted/50 flex items-center justify-center">
          <Minus className="h-3 w-3 text-muted-foreground/50" />
        </div>
      </div>
    );
  }
  return (
    <div className="text-center" data-testid="feature-partial">
      <span className="text-xs font-medium text-foreground/80 bg-primary/5 border border-primary/10 rounded-md px-2 py-0.5 inline-block">
        {value}
      </span>
    </div>
  );
}

export default function LicenseComparisonPage() {
  const [, navigate] = useLocation();
  const searchString = useSearch();
  const params = new URLSearchParams(searchString);

  const compareParam = params.get("compare");
  const [selectedIds, setSelectedIds] = useState<string[]>(() => {
    if (compareParam) {
      const name = decodeURIComponent(compareParam);
      const license = LICENSES.find((l) => l.displayName === name);
      return license ? [license.skuPartNumber] : [];
    }
    return [];
  });

  useEffect(() => {
    if (compareParam) {
      const name = decodeURIComponent(compareParam);
      const license = LICENSES.find((l) => l.displayName === name);
      if (license && !selectedIds.includes(license.skuPartNumber)) {
        setSelectedIds([license.skuPartNumber]);
      }
    }
  }, [searchString]);

  const selectedLicenses = useMemo(() => {
    return selectedIds
      .map((id) => LICENSES.find((l) => l.skuPartNumber === id))
      .filter(Boolean) as LicenseInfo[];
  }, [selectedIds]);

  const handleSelect = (index: number, skuPartNumber: string) => {
    setSelectedIds((prev) => {
      const next = [...prev];
      if (skuPartNumber === EMPTY) {
        next.splice(index, 1);
      } else if (index < next.length) {
        next[index] = skuPartNumber;
      } else {
        next.push(skuPartNumber);
      }
      return next;
    });
  };

  const handleRemove = (index: number) => {
    setSelectedIds((prev) => prev.filter((_, i) => i !== index));
  };

  const handleClearAll = () => setSelectedIds([]);

  const canAdd = selectedIds.length < 3;

  const visibleCategories = useMemo(() => {
    if (selectedLicenses.length === 0) return [];
    return FEATURE_CATEGORIES.filter((cat) =>
      cat.features.some((f) =>
        selectedLicenses.some((l) => l.features[f.key] !== false)
      )
    );
  }, [selectedLicenses]);

  const suiteLicenses = LICENSES.filter((l) => l.category === "Suite");
  const addonLicenses = LICENSES.filter((l) => l.category !== "Suite");

  return (
    <div className="min-h-screen bg-background flex flex-col font-sans text-foreground">
      <header className="sticky top-0 z-10 bg-card/80 backdrop-blur-md border-b border-border px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div
            className="flex items-center gap-2 cursor-pointer"
            onClick={() => navigate("/")}
            data-testid="link-home"
          >
            <div className="h-8 w-8 rounded-md bg-primary flex items-center justify-center text-primary-foreground font-bold">
              A
            </div>
            <h1 className="text-xl font-semibold tracking-tight">Astra</h1>
          </div>
          <nav className="hidden sm:flex items-center gap-1 ml-4">
            <Button
              variant="ghost"
              size="sm"
              className="text-muted-foreground hover:text-foreground"
              onClick={() => navigate("/")}
              data-testid="nav-dashboard"
            >
              Dashboard
            </Button>
            <Button
              variant="ghost"
              size="sm"
              className="text-foreground font-medium bg-muted/50"
              data-testid="nav-licenses"
            >
              License Guide
            </Button>
          </nav>
        </div>
      </header>

      <main className="flex-1 p-8 max-w-7xl mx-auto w-full space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
        <div className="flex flex-col gap-2">
          <div className="flex items-center gap-3">
            <Button
              variant="ghost"
              size="icon"
              className="h-8 w-8 sm:hidden"
              onClick={() => navigate("/")}
              data-testid="button-back"
            >
              <ArrowLeft className="h-4 w-4" />
            </Button>
            <div>
              <h2 className="text-3xl font-display font-semibold">
                License Comparison Guide
              </h2>
              <p className="text-muted-foreground">
                Compare up to 3 Microsoft 365 licenses side by side. Feature
                data reflects current Microsoft documentation.
              </p>
            </div>
          </div>
        </div>

        <Card className="border-border/50 shadow-sm">
          <CardHeader className="pb-3">
            <div className="flex items-center justify-between">
              <div>
                <CardTitle className="text-base">Select Licenses</CardTitle>
                <CardDescription>
                  Choose up to 3 licenses to compare features, pricing, and
                  capabilities.
                </CardDescription>
              </div>
              {selectedIds.length > 0 && (
                <Button
                  variant="ghost"
                  size="sm"
                  className="text-muted-foreground text-xs"
                  onClick={handleClearAll}
                  data-testid="button-clear-all"
                >
                  <Trash2 className="h-3 w-3 mr-1" />
                  Clear All
                </Button>
              )}
            </div>
          </CardHeader>
          <CardContent>
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              {Array.from({ length: Math.max(selectedIds.length + (canAdd ? 1 : 0), 1) }).map(
                (_, index) => {
                  const isExisting = index < selectedIds.length;
                  const currentSku = isExisting ? selectedIds[index] : EMPTY;
                  const alreadySelected = new Set(selectedIds);

                  return (
                    <div key={index} className="space-y-2">
                      <div className="flex items-center gap-2">
                        <Select
                          value={currentSku}
                          onValueChange={(v) => handleSelect(index, v)}
                        >
                          <SelectTrigger
                            className="bg-background"
                            data-testid={`select-license-${index}`}
                          >
                            <SelectValue placeholder="Select a license..." />
                          </SelectTrigger>
                          <SelectContent>
                            {isExisting && (
                              <SelectItem value={EMPTY}>
                                <span className="text-muted-foreground">
                                  Remove selection
                                </span>
                              </SelectItem>
                            )}
                            <div className="px-2 py-1.5 text-xs font-semibold text-muted-foreground">
                              Suite Licenses
                            </div>
                            {suiteLicenses.map((l) => (
                              <SelectItem
                                key={l.skuPartNumber}
                                value={l.skuPartNumber}
                                disabled={
                                  alreadySelected.has(l.skuPartNumber) &&
                                  currentSku !== l.skuPartNumber
                                }
                              >
                                <div className="flex items-center justify-between w-full gap-2">
                                  <span>{l.displayName}</span>
                                  <span className="text-muted-foreground text-xs">
                                    ${l.costPerMonth}/mo
                                  </span>
                                </div>
                              </SelectItem>
                            ))}
                            <div className="px-2 py-1.5 text-xs font-semibold text-muted-foreground mt-1">
                              Add-ons & Standalone
                            </div>
                            {addonLicenses.map((l) => (
                              <SelectItem
                                key={l.skuPartNumber}
                                value={l.skuPartNumber}
                                disabled={
                                  alreadySelected.has(l.skuPartNumber) &&
                                  currentSku !== l.skuPartNumber
                                }
                              >
                                <div className="flex items-center justify-between w-full gap-2">
                                  <span>{l.displayName}</span>
                                  <span className="text-muted-foreground text-xs">
                                    ${l.costPerMonth}/mo
                                  </span>
                                </div>
                              </SelectItem>
                            ))}
                          </SelectContent>
                        </Select>
                        {isExisting && (
                          <Button
                            variant="ghost"
                            size="icon"
                            className="h-9 w-9 shrink-0 text-muted-foreground hover:text-destructive"
                            onClick={() => handleRemove(index)}
                            data-testid={`button-remove-license-${index}`}
                          >
                            <XIcon className="h-4 w-4" />
                          </Button>
                        )}
                      </div>
                      {isExisting && selectedLicenses[index] && (
                        <div className="rounded-lg border border-border/50 p-3 bg-muted/10">
                          <div className="flex items-center gap-2 mb-1">
                            <Badge
                              variant="outline"
                              className={`text-[10px] ${selectedLicenses[index].category === "Suite" ? "border-blue-500/30 text-blue-600 bg-blue-500/5" : "border-amber-500/30 text-amber-600 bg-amber-500/5"}`}
                            >
                              {selectedLicenses[index].category}
                            </Badge>
                            <span className="text-sm font-semibold text-primary">
                              ${selectedLicenses[index].costPerMonth}/mo
                            </span>
                          </div>
                          <p className="text-xs text-muted-foreground leading-relaxed">
                            {selectedLicenses[index].description}
                          </p>
                        </div>
                      )}
                    </div>
                  );
                }
              )}
            </div>
          </CardContent>
        </Card>

        {selectedLicenses.length === 0 ? (
          <Card className="border-border/50 shadow-sm">
            <CardContent className="py-16 text-center">
              <div className="flex flex-col items-center gap-3">
                <div className="h-12 w-12 rounded-full bg-muted/50 flex items-center justify-center">
                  <Plus className="h-6 w-6 text-muted-foreground" />
                </div>
                <div>
                  <p className="font-medium text-muted-foreground">
                    Select licenses above to compare
                  </p>
                  <p className="text-sm text-muted-foreground/70 mt-1">
                    Choose 1 to 3 licenses to see a detailed feature-by-feature
                    comparison
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        ) : (
          <Card className="border-border/50 shadow-sm overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="border-b border-border bg-muted/30">
                    <th className="text-left text-sm font-medium text-muted-foreground px-6 py-3 w-[280px] min-w-[200px]">
                      Feature
                    </th>
                    {selectedLicenses.map((license) => (
                      <th
                        key={license.skuPartNumber}
                        className="text-center px-4 py-3 min-w-[180px]"
                      >
                        <div className="space-y-1">
                          <div className="font-semibold text-sm">
                            {license.displayName}
                          </div>
                          <div className="text-primary font-bold text-lg">
                            ${license.costPerMonth}
                            <span className="text-xs font-normal text-muted-foreground">
                              /user/mo
                            </span>
                          </div>
                        </div>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleCategories.map((category) => (
                    <Fragment key={category.name}>
                      <tr
                        className="bg-muted/10 border-y border-border/50"
                      >
                        <td
                          colSpan={selectedLicenses.length + 1}
                          className="px-6 py-2"
                        >
                          <span className="text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                            {category.name}
                          </span>
                        </td>
                      </tr>
                      {category.features.map((feature) => {
                        const values = selectedLicenses.map(
                          (l) => l.features[feature.key]
                        );
                        const allFalse = values.every((v) => v === false);
                        if (allFalse) return null;

                        const allSame =
                          values.length > 1 &&
                          values.every((v) => JSON.stringify(v) === JSON.stringify(values[0]));

                        return (
                          <tr
                            key={feature.key}
                            className="border-b border-border/30 hover:bg-muted/10 transition-colors"
                            data-testid={`row-feature-${feature.key}`}
                          >
                            <td className="px-6 py-2.5 text-sm text-foreground/80">
                              {feature.name}
                            </td>
                            {selectedLicenses.map((license) => (
                              <td
                                key={license.skuPartNumber}
                                className={`px-4 py-2.5 ${allSame ? "" : ""}`}
                              >
                                {renderFeatureValue(
                                  license.features[feature.key]
                                )}
                              </td>
                            ))}
                          </tr>
                        );
                      })}
                    </Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        )}

        <div className="text-center text-xs text-muted-foreground pb-4">
          <p>
            Feature data is based on publicly available Microsoft 365
            documentation and may not reflect the latest changes.
          </p>
          <p className="mt-1">
            <a
              href="https://www.microsoft.com/en-us/microsoft-365/business/compare-all-plans"
              target="_blank"
              rel="noopener noreferrer"
              className="text-primary hover:underline inline-flex items-center gap-1"
              data-testid="link-microsoft-plans"
            >
              View official Microsoft 365 plan comparison
              <ExternalLink className="h-3 w-3" />
            </a>
          </p>
        </div>
      </main>

      <footer className="py-4 text-center text-xs text-muted-foreground border-t border-border/30">
        &copy; 2026 Cavaridge, LLC. All rights reserved.
      </footer>
    </div>
  );
}
