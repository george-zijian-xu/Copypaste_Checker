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
      title: "Instantly find plagiarism by pasting or uploading your research paper, essay, or article.",
    },
    {
      icon: FileCheck,
      title: "Quickly ensure integrity by checking your work against billions of web pages with one click.",
    },
    {
      icon: Lightbulb,
      title: "Get an originality score for your document to see how unique your ideas are.",
    },
    {
      icon: Speedometer,
      title:
        "Speed up work with recommendations on what—and how—to cite, as well as real-time feedback on your writing.",
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