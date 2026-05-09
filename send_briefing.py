      - name: Send briefing
        env:
          MS_TENANT_ID: ${{ secrets.MS_TENANT_ID }}
          MS_CLIENT_ID: ${{ secrets.MS_CLIENT_ID }}
          MS_CLIENT_SECRET: ${{ secrets.MS_CLIENT_SECRET }}
          MS_USER_EMAIL: ${{ secrets.MS_USER_EMAIL }}
        run: python send_briefing.py
