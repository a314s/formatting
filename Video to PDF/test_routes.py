#!/usr/bin/env python
"""
Test script to verify that all routes in the application are working correctly.
This script will make requests to each route and report any issues.
"""

import os
import sys
import requests
from urllib.parse import urljoin

def test_routes(base_url):
    """Test all routes in the application"""
    print(f"Testing routes at {base_url}")
    
    # List of routes to test
    routes = [
        "/",                  # Main page
        "/settings",          # Settings page
        "/get_frames",        # Get frames page
        "/convert_to_pdf"     # Convert to PDF page
    ]
    
    # Test each route
    for route in routes:
        url = urljoin(base_url, route)
        print(f"\nTesting route: {url}")
        
        try:
            response = requests.get(url, timeout=10)
            
            if response.status_code == 200:
                print(f"✅ SUCCESS: Route {route} is working (Status: {response.status_code})")
            else:
                print(f"❌ ERROR: Route {route} returned status code {response.status_code}")
                print(f"Response: {response.text[:200]}...")
        except requests.RequestException as e:
            print(f"❌ ERROR: Could not connect to {url}")
            print(f"Exception: {str(e)}")
    
    print("\nRoute testing completed.")
    print("If any routes failed, check the following:")
    print("1. Make sure the application is running")
    print("2. Verify that all template files are in the correct location")
    print("3. Check the application logs for any errors")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        base_url = sys.argv[1]
    else:
        # Default to localhost if no URL is provided
        base_url = "http://localhost:5000"
        print("No base URL provided, using default: http://localhost:5000")
        print("To specify a different URL, run: python test_routes.py YOUR_URL")
    
    test_routes(base_url)