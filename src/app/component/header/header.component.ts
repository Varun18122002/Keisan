import { Component } from '@angular/core';

@Component({
  selector: 'app-header',
  template: `
    <div class="header">
      <div class="header-left">
        <div class="logo">
          <span class="logo-icon">📊</span>
          <span class="logo-text">Keisan</span>
        </div>
      </div>
  
    </div>
  `,
  styles: `
    .header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 8px 16px;
      background-color: #f8f9fa;
      border-bottom: 1px solid #d3d3d3;
      font-family: Arial, sans-serif;
    }

    .header-left {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .logo {
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 18px;
      font-weight: bold;
      color: #333;
    }

    .logo-icon {
      font-size: 20px;
      color: #4caf50; /* Google green color */
    }

    .logo-text {
      font-size: 18px;
      color: #333;
    }

    .header-right {
      display: flex;
      align-items: center;
      gap: 12px;
    }

    .action-btn {
      padding: 6px 12px;
      border: 1px solid #d3d3d3;
      border-radius: 4px;
      background-color: #fff;
      color: #333;
      cursor: pointer;
      font-size: 14px;
      font-weight: bold;
      transition: background-color 0.2s, box-shadow 0.2s;
    }

    .action-btn:hover {
      background-color: #f1f1f1;
      box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }

    .profile-pic {
      width: 32px;
      height: 32px;
      border-radius: 50%;
      border: 1px solid #d3d3d3;
      cursor: pointer;
    }
  `,
})
export class HeaderComponent {}
