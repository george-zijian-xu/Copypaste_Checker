import { Card, CardContent } from "@/components/ui/card";
import { Lightbulb, FileCheck, ScanSearch, GaugeIcon as Speedometer } from "lucide-react";
import type React from "react";

interface FeatureCardProps {
  icon: React.ElementType;
  title: string;
}

const FeatureCard: React.FC<FeatureCardProps> = ({ icon: Icon, title }) => (
  <Card className="text-center p-6 border-none shadow-none bg-transparent">
    <CardContent className="flex flex-col items-center p-0 space-y-4">
      <div className="p-3 rounded-full bg-green-010 text-green-060">
        <Icon className="h-8 w-8" />
      </div>
      <h3 className="text-lg font-semibold text-gray-900">{title}</h3>
    </CardContent>
  </Card>
);

export default function CardsSection() {
  const features = [
    {
      icon: ScanSearch,
      title: "Sharp-Detection: Identify every one-shot paste at the character level using the document’s internal RSID timestamps.",
    },
    {
      icon: FileCheck,
      title: "Noise-Less: Automatically filter out simple formatting or style edits, leaving pure paste events.",
    },
    {
      icon: Lightbulb,
      title: "Audit-Friendly: Export raw XML snippets plus plain-English summaries for each highlight, perfect for appeals and compliance.",
    },
    {
      icon: Speedometer,
      title:
        "API-Ready: Integrate seamlessly to scan batches of documents, generate reports, and enforce integrity at scale.",
    },
  ];

  return (
    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-8">
      {features.map((feature, index) => (
        <FeatureCard key={index} icon={feature.icon} title={feature.title} />
      ))}
    </div>
  );
} 