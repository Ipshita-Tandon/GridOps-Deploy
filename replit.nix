{ pkgs }: {
  deps = [
    pkgs.python3
    pkgs.nodejs-18_x
    pkgs.nodePackages.npm
    pkgs.ngrok
  ];
} 