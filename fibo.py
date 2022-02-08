def Fibo(n):
  if n == 0: return 1
  produit = 1
  for i in range(1, n+1):
    produit = produit * i
  return produit
 
  
