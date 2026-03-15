SELECT 
  c.email,
  l.template_id,
  l.private_key,
  w.balance
FROM customers c
JOIN licenses l ON l.customer_id = c.id
JOIN credit_wallets w ON w.license_id = l.id
WHERE c.email = 'test@example.com';         // Insert Customer Email here
