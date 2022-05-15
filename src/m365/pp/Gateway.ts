export interface Gateway { 
  id: string; 
  gatewayId: string; 
  name: string; 
  type: string; 
  publicKey: { exponent: string; 
    modulus: string 
  }; 
  gatewayAnnotation: string 
}