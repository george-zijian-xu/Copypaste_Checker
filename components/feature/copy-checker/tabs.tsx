"use client";

import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import ImagePanel from "./panels";

export default function TabsSection() {
  return (
    <Tabs defaultValue="plagiarism-ai" className="w-full">
      <TabsList className="grid w-full grid-cols-5 h-auto bg-transparent border-b border-gray-200 rounded-none p-0">
        {[
          { value: "plagiarism-ai", label: "Check plagiarism & AI" },
          { value: "authentic-authorship", label: "Authentic authorship" },
          { value: "citation-suggestions", label: "Citation suggestions" },
          { value: "perfect-proofreading", label: "Perfect proofreading" },
          { value: "ai-integrity", label: "AI integrity" },
        ].map((tab) => (
          <TabsTrigger
            key={tab.value}
            value={tab.value}
            className="data-[state=active]:bg-transparent data-[state=active]:shadow-none data-[state=active]:border-b-2 data-[state=active]:border-green-050 data-[state=active]:text-green-050 text-gray-700 font-semibold py-4 rounded-none"
          >
            {tab.label}
          </TabsTrigger>
        ))}
      </TabsList>

      <TabsContent value="plagiarism-ai" className="pt-8 text-center text-gray-600">
        <p>
          Our plagiarism checker helps you ensure the originality of your work by comparing it against billions of web
          pages and academic papers. It highlights passages that may need citations and provides suggestions for proper
          attribution.
        </p>
      </TabsContent>

      <TabsContent value="authentic-authorship" className="pt-8">
        <ImagePanel
          imageSrc="/placeholder.svg?height=300&width=400"
          imageAlt="Authentic authorship example"
          textBoxTitle="Bring transparency to your work"
          textBoxContent="Using a variety of text sources in your content? Grammarly Authorship automatically categorizes your text based on where it came from (generative AI, an online database, typed by you, etc.) so that you can easily show your work and confidently submit your most original writing."
          buttonText="Learn More"
        />
      </TabsContent>

      <TabsContent value="citation-suggestions" className="pt-8 text-center text-gray-600">
        <p>
          Get smart suggestions for citations in various styles (APA, MLA, Chicago, etc.). Our checker helps you
          correctly attribute sources and avoid unintentional plagiarism.
        </p>
      </TabsContent>

      <TabsContent value="perfect-proofreading" className="pt-8 text-center text-gray-600">
        <p>
          Beyond plagiarism, our advanced proofreading capabilities catch grammar, spelling, punctuation, and style
          errors, ensuring your document is polished and professional.
        </p>
      </TabsContent>

      <TabsContent value="ai-integrity" className="pt-8">
        <ImagePanel
          imageSrc="/images/ai-integrity-circuit.png"
          imageAlt="AI integrity example showing AI generated text detection"
          textBoxTitle="Acknowledge Grammarly gen AI use"
          textBoxContent="Grammarly was used to create an outline for this business pitch using the prompt 'Draft an outline'. Additionally, it was used to rewrite parts of the pitch with the prompts 'Shorten it' and 'Make it more persuasive'."
          buttonText="Insert Code"
        />
      </TabsContent>
    </Tabs>
  );
} 