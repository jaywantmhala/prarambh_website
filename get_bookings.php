<?php
// get_bookings.php

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET");
header("Content-Type: application/json; charset=UTF-8");

$host = "localhost";
$username = "root";
$password = "root"; // Matching updated MySQL password
$dbname = "kiran";

// Create connection to MySQL
$conn = new mysqli($host, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    die(json_encode(["status" => "error", "message" => "MySQL Connection failed: " . $conn->connect_error]));
}

// Fetch all bookings ordered by latest signup
$sql = "SELECT `id`, `name`, `mobile`, `city`, `payment_status`, `created_at` FROM `bookings` ORDER BY `created_at` DESC";
$result = $conn->query($sql);

$bookings = [];

if ($result && $result->num_rows > 0) {
    while($row = $result->fetch_assoc()) {
        $bookings[] = $row;
    }
}

echo json_encode(["status" => "success", "bookings" => $bookings]);

$conn->close();
?>
