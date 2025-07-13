import Link from "next/link";
import { BookText } from "lucide-react";

export function Header() {
  return (
    <header className="w-full bg-white border-b border-gray-200 py-4 px-4 sm:px-6 lg:px-8 flex items-center justify-between">
      <Link href="#" className="flex items-center space-x-2 text-gray-900 font-bold text-xl" prefetch={false}>
        <BookText className="h-6 w-6 text-green-050" />
        <span>Copy Forensics</span>
      </Link>
      <nav>
        <ul className="flex space-x-4">
          {/* <li>
            <Link href="#" className="text-gray-600 hover:text-green-050 transition-colors" prefetch={false}>
              Features
            </Link>
          </li>
          <li>
            <Link href="#" className="text-gray-600 hover:text-green-050 transition-colors" prefetch={false}>
              Pricing
            </Link>
          </li>
          <li>
            <Link href="#" className="text-gray-600 hover:text-green-050 transition-colors" prefetch={false}>
              Contact
            </Link>
          </li> */}
        </ul>
      </nav>
    </header>
  );
} 