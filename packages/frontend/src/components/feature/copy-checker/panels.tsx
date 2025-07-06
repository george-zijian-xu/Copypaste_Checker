import Image from "next/image";

interface PanelProps {
  imageSrc: string;
  imageAlt: string;
  textBoxTitle: string;
  textBoxContent: string;
  buttonText?: string;
}

export default function ImagePanel({
  imageSrc,
  imageAlt,
  textBoxTitle,
  textBoxContent,
  buttonText,
}: PanelProps) {
  return (
    <div className="grid grid-cols-1 md:grid-cols-2 gap-8 items-center pt-8">
      <div className="relative w-full h-auto flex justify-center">
        <Image
          src={imageSrc || "/placeholder.svg"}
          alt={imageAlt}
          width={400}
          height={300}
          layout="responsive"
          objectFit="contain"
          className="rounded-lg shadow-lg"
        />
      </div>
      <div className="relative p-6 bg-white border border-gray-200 rounded-lg shadow-md">
        <div className="absolute -top-3 -left-3 bg-green-050 text-white px-3 py-1 rounded-md text-sm font-semibold">
          {textBoxTitle}
        </div>
        <p className="text-sm text-gray-700 mt-2">{textBoxContent}</p>
        {buttonText && (
          <button className="mt-4 inline-flex items-center justify-center rounded-md text-sm font-medium ring-offset-background transition-colors hover:bg-green-060 h-10 px-4 py-2 bg-green-050 text-white">
            {buttonText}
          </button>
        )}
      </div>
    </div>
  );
}
