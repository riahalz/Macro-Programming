function PRODUCT_DISCOUNT(price, discountPercentage) {
  // Check if inputs are numbers
  if (typeof price !== 'number' || typeof discountPercentage !== 'number'){
    return "Error: Inputs must be numbers";
  }

  // Check if discountPercentage is within valid range
  if (discountPercentage < 0 || discountPercentage > 100){
    return "Error: Discount must be between 0 and 100";
  }
  
  // Calculate discount amount
  var discountAmount = price * (discountPercentage / 100);

  // Return the final price after discount
  return price - discountAmount;
}
