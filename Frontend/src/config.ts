// Environment configuration
interface Config {
  apiUrl: string;
  wopiUrl: string;
}

// Development configuration
const devConfig: Config = {
  apiUrl: 'http://localhost:5000',
  wopiUrl: 'http://localhost:3001'
};

// Production configuration
const prodConfig: Config = {
  apiUrl: import.meta.env.VITE_API_URL || 'http://10.172.168.165:5000',
  wopiUrl: import.meta.env.VITE_WOPI_URL || 'http://localhost:3001'
};

// Select configuration based on environment
const config: Config = import.meta.env.PROD ? prodConfig : devConfig;

export default config; 