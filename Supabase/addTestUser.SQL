## Step 1 ##
INSERT INTO customers (username, email)
VALUES ('Test User', 'test@example.com');

## Step 2 ##
INSERT INTO licenses (customer_id, template_id, private_key)
VALUES (
  (SELECT id FROM customers WHERE email = 'test@example.com'),
  'internship-tracker',
  'TEST-KEY-12345'
);

## Step 3 ##
INSERT INTO credit_wallets (license_id, balance)
VALUES (
  (SELECT id FROM licenses WHERE private_key = 'TEST-KEY-12345'),
  10  -- give them 10 test credits
);