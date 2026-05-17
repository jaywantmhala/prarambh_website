<?php
// save_booking.php

// Enable error reporting to find issues
error_reporting(E_ALL);
ini_set('display_errors', 1);

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: POST");
header("Content-Type: application/json; charset=UTF-8");

// Database configuration
$host = "localhost";
$username = "root"; // Update this if your database username is different
$password = "root"; // Update this if your database password is different
$dbname = "kiran";

// Create connection to MySQL
$conn = new mysqli($host, $username, $password);

// Check connection
if ($conn->connect_error) {
    die(json_encode(["status" => "error", "message" => "MySQL Connection failed: " . $conn->connect_error]));
}

// 1. Create database if not exists
$sql = "CREATE DATABASE IF NOT EXISTS `$dbname`";
if ($conn->query($sql) !== TRUE) {
    die(json_encode(["status" => "error", "message" => "Error creating database: " . $conn->error]));
}

// 2. Select the database
$conn->select_db($dbname);

// 3. Create table if not exists
$tableSql = "CREATE TABLE IF NOT EXISTS `bookings` (
    `id` INT(6) UNSIGNED AUTO_INCREMENT PRIMARY KEY,
    `name` VARCHAR(100) NOT NULL,
    `mobile` VARCHAR(15) NOT NULL,
    `city` VARCHAR(100) NOT NULL,
    `payment_status` VARCHAR(50) DEFAULT 'Pending',
    `created_at` TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)";

if ($conn->query($tableSql) !== TRUE) {
    die(json_encode(["status" => "error", "message" => "Error creating table: " . $conn->error]));
}

// Get POST data sent from Javascript fetch()
$data = json_decode(file_get_contents("php://input"));

// If someone just opens this PHP file in the browser directly to test
if ($data === null) {
    echo json_encode(["status" => "success", "message" => "Database '$dbname' and table 'bookings' have been successfully created! Everything is working."]);
    $conn->close();
    exit();
}

// 4. Save the lead into the database
if (isset($data->name) && isset($data->mobile) && isset($data->city)) {
    
    // Prevent SQL Injection
    $name = $conn->real_escape_string($data->name);
    $mobile = $conn->real_escape_string($data->mobile);
    $city = $conn->real_escape_string($data->city);
    $status = "Pending/Initiated";
    
    $stmt = $conn->prepare("INSERT INTO `bookings` (`name`, `mobile`, `city`, `payment_status`) VALUES (?, ?, ?, ?)");
    
    if (!$stmt) {
        die(json_encode(["status" => "error", "message" => "Statement prepare failed: " . $conn->error]));
    }

    $stmt->bind_param("ssss", $name, $mobile, $city, $status);
    
    if ($stmt->execute()) {
        echo json_encode(["status" => "success", "message" => "Booking saved successfully"]);
    } else {
        echo json_encode(["status" => "error", "message" => "Error inserting record: " . $stmt->error]);
    }
    
    $stmt->close();
} else {
    echo json_encode(["status" => "error", "message" => "Incomplete data provided"]);
}

$conn->close();
?>
