const LEGACY_VENDOR_STYLESHEETS = [
  'https://unpkg.com/boxicons@2.1.4/css/boxicons.min.css',
  'https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/4.1.3/css/bootstrap.min.css',
  'https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200',
  'https://fonts.googleapis.com/css?family=Permanent+Marker',
  'https://fonts.googleapis.com/css2?family=Noto+Sans&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css',
];

export function legacyPageStylesheets(...pageStylesheets: string[]): string[] {
  return ['/legacy/style.css', '/legacy/banner.css', ...pageStylesheets, ...LEGACY_VENDOR_STYLESHEETS];
}
