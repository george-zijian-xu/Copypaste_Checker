# Deployment Guide: Vercel + Railway

This guide will help you deploy your Copy-Paste Checker application using Vercel for the frontend and Railway for the backend.

## 🚀 **Backend Deployment to Railway**

### Step 1: Prepare Your Repository
1. Make sure your changes are committed and pushed to GitHub
2. Ensure your `packages/backend/` directory has all the necessary files

### Step 2: Deploy to Railway
1. Go to [Railway](https://railway.app) and sign up/login
2. Click "New Project" → "Deploy from GitHub repo"
3. Select your repository
4. Choose "Deploy from subdirectory" and set it to `packages/backend`
5. Railway will automatically detect your Python app

### Step 3: Configure Environment Variables
In your Railway project dashboard, go to Variables and add:
```
ENVIRONMENT=production
```

### Step 4: Get Your Railway URL
After deployment, Railway will provide you with a URL like:
`https://your-app-name.railway.app`

---

## 🌐 **Frontend Deployment to Vercel**

### Step 1: Deploy to Vercel
1. Go to [Vercel](https://vercel.com) and sign up/login
2. Click "New Project" → Import from GitHub
3. Select your repository
4. Set the **Root Directory** to `packages/frontend`
5. Vercel will automatically detect your Next.js app

### Step 2: Configure Environment Variables
In your Vercel project dashboard, go to Settings → Environment Variables and add:
```
NEXT_PUBLIC_API_URL=https://your-railway-app.railway.app
```
(Replace with your actual Railway URL from Step 4 above)

### Step 3: Update CORS Configuration
After you get your Vercel URL, update `packages/backend/src/middleware/cors_middleware.py`:
```python
origins = [
    "http://localhost",
    "http://localhost:3000",
    "http://localhost:3001",
    "https://your-vercel-app.vercel.app",  # Replace with your actual Vercel URL
    "https://*.vercel.app",
]
```

### Step 4: Redeploy Backend
Commit and push the CORS changes to trigger a new Railway deployment.

---

## 🔧 **Local Development**

### Backend
```bash
cd packages/backend
pip install -r requirements.txt
uvicorn server:app --reload --port 8000
```

### Frontend
```bash
cd packages/frontend
npm install
npm run dev
```

Make sure to create a `.env.local` file in `packages/frontend/` with:
```
NEXT_PUBLIC_API_URL=http://localhost:8000
```

---

## 📋 **Deployment Checklist**

- [ ] Backend deployed to Railway
- [ ] Frontend deployed to Vercel
- [ ] Environment variables configured on both platforms
- [ ] CORS origins updated with production URLs
- [ ] Both applications can communicate successfully

---

## 🐛 **Troubleshooting**

### Common Issues

1. **CORS Errors**: Make sure your Vercel URL is added to the CORS origins list
2. **API Connection Issues**: Verify the `NEXT_PUBLIC_API_URL` environment variable is correct
3. **Build Failures**: Check that all dependencies are listed in `requirements.txt` and `package.json`

### Logs
- **Railway**: Check logs in your Railway dashboard
- **Vercel**: Check logs in your Vercel dashboard under the "Functions" tab

---

## 📚 **Additional Resources**

- [Railway Documentation](https://docs.railway.app)
- [Vercel Documentation](https://vercel.com/docs)
- [Next.js Deployment Guide](https://nextjs.org/docs/deployment)
- [FastAPI Deployment Guide](https://fastapi.tiangolo.com/deployment/) 