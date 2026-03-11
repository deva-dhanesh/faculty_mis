module.exports = {
  apps: [
    {
      name: "faculty_mis",
      script: "app.py",
      interpreter: "venv/bin/python",
      cwd: "/home/fet/all-projects/faculty_mis",
      env: {
        PORT: 6201,
        HOST: "0.0.0.0"
      }
    }
  ]
};
