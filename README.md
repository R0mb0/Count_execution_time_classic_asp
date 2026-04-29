# Count execution time in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/e407e0644e734635811fe1cad5d19409)](https://app.codacy.com/gh/R0mb0/Count_execution_time_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Count_execution_time_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Count_execution_time_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

---

<div align="center">
  <a href="http://paypal.me/R0mb0">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://github.com/R0mb0/Support_the_dev_badge/blob/main/Badge/SVG/Support_the_dev_badge_Dark.svg">
      <source media="(prefers-color-scheme: light)" srcset="https://github.com/R0mb0/Support_the_dev_badge/blob/main/Badge/SVG/Support_the_dev_badge_Light.svg">
      <img alt="Saved you time? Support the dev" src="https://github.com/R0mb0/Support_the_dev_badge/blob/main/Badge/SVG/Support_the_dev_badge_Default.svg">
    </picture>
  </a>
</div>


## `count_execution_time.class.asp`'s avaible functions

- Set the starting time for counting execution time -> `Public Function set_start_time()`
- Retrieve the execution time at the moment -> `Public Function get_execution_time()`
  >
  > The time is retrieved in seconds

## How to use 

> From `Test.asp`

1. Initialize the class
   ```asp
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="count_execution_time.class.asp"-->
   <%
      Dim time
      Set time = new count_execution_time
   ```
2. Set the starting time
   ```asp
   time.set_start_time()
   ```
3. Retrieve the execution time
   ```asp
     Response.write(time.get_execution_time())
   %>
   ```
   
 <a href="https://github.com/R0mb0/Not_made_by_AI">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDark.svg">
    <source media="(prefers-color-scheme: light)" srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAILight.svg">
    <img alt="Not made by AI" src="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDefault.svg">
  </picture>
</a>
