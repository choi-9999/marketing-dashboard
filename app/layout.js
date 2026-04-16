import "./globals.css";

export const metadata = {
  title: "지점 활성화 방안 대시보드",
  description: "지점 활성화 방안 대시보드와 RAWDATA 편집 화면",
  icons: {
    icon: "/favicon.png"
  }
};

export default function RootLayout({ children }) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
