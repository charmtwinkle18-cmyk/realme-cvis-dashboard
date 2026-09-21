import "./globals.css";

export const metadata = {
  title: "realme CVIS Sales Dashboard",
  description: "CVIS sales intelligence dashboard powered by Excel uploads",
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
