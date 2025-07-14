export function Footer() {
  return (
    <footer className="w-full bg-gray-100 py-8 px-4 sm:px-6 lg:px-8 text-center text-gray-600 text-sm">
      <div className="max-w-6xl mx-auto space-y-4">
        <div className="flex flex-wrap justify-center gap-x-8 gap-y-2">
          {/* <a href="#" className="hover:underline">
            Privacy Policy
          </a>
          <a href="#" className="hover:underline">
            Terms of Service
          </a>
          <a href="#" className="hover:underline">
            Support
          </a>
          <a href="#" className="hover:underline">
            About Us
          </a> */}
        </div>
        <p>&copy; {new Date().getFullYear()} IsItYours. All rights reserved.</p>
      </div>
    </footer>
  );
} 