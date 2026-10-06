import './globals.css';

export const metadata = {
  title: 'Student Search',
  description: 'Search students from an uploaded spreadsheet',
};

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
