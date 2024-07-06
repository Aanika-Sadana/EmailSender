### Automatically send emails from .CSV files exported from CrunchBase

Instructions:
1. Install Visual Studio Code for your operating system: https://code.visualstudio.com/download
2. Open Visual Studio Code.
3. Click the **_Extensions_** button on the left sidebar and search for "Python". Install and enable the **_Python_** extension.
  <img width="250" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/b461a173-d2e2-4442-885b-64c12a4adf6c">
  <img width="300" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/5c63c7ae-9a90-441d-ad71-a86ffb62a653">

4. Navigate to **_Explorer_** on the left sidebar and select **_Clone Repository_**.
  <img width="200" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/09e1b68f-95b7-471d-a94a-58b4d21600d5">
  <img width="225" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/55dcdef8-0121-4a22-b81f-804bd707cbfa">

5. Enter this repository URL in the text box at the top: https://github.com/Aanika-Sadana/EmailSender.git
  <img width="500" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/ad2464f4-115a-4eb7-8193-b6b0c3f6084b">

6. Select a folder to save the repository in and click **_Select as Repository Destination_**.
  <img width="380" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/473ee245-26bc-412d-a544-42adc2567b6d">

7. Your workspace should contain the following files:
- **EMAILSENDER**  _(Workspace)_
  - **automation**  _(directory containing Python script)_
    - **email_sender.py**  _(Python script)_
  - **batch 1.1**  _(Batch 1.1)_
  - **batch 1.2**  _(Batch 1.2)_
  - **batch ...**  _(Batch ...)_
  - **README.md**  _(instructions)_

  <img width="300" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/06fc97b6-cf36-45c5-b69f-792a7e2ef165">

Your script is now set up in VS Code. Refer to steps 8-10 for every email batch exported from CrunchBase.

8. To upload a new .CSV file generated by CrunchBase, drag and drop your .CSV file from your file explorer to VS Code
  <img width="450" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/c4c4a998-bf3b-4b3f-ad8e-90757b7b8e6d">
  <img width="350" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/9fe68004-8f76-4652-93d1-75fd95e7d962">

9. Open the script called "email_sender.py" and replace the file name with the name of the newly uploaded .CSV file
  <img width="500" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/dde46754-a699-481f-bbe9-a851cfbadbb7">
  <img width="350" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/95ce8922-fdfd-41ac-b295-aa9d69912f88">

10. Hit **_Run Python File_** on the top right to run the script and automatically send all emails listed in the batch.
    
  <img width="300" alt="image" src="https://github.com/Aanika-Sadana/EmailSender/assets/70586980/420a80c1-bc8b-4707-8884-32c3d0514768">


    
