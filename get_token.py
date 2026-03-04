"""
Cafe24 OAuth2 Access Token 발급 스크립트
1) 브라우저에서 인증
2) 리다이렉트된 URL을 복사해서 붙여넣기
3) access_token, refresh_token 저장
"""
import webbrowser
import requests
import base64
import json
import urllib.parse

# ── 설정 ──
MALL_ID = "rkghrud1"
CLIENT_ID = "CfoSdFlVT6oDL1VUWN9jwB"
CLIENT_SECRET = "GNXM72AhBReZjZ3GIMEyfO"
REDIRECT_URI = "https://e356-14-6-135-151.ngrok-free.app/oauth/callback"

# 출고/송장 관리에 필요한 scope
SCOPES = ",".join([
    "mall.read_order",
    "mall.read_shipping",
    "mall.write_shipping",
    "mall.read_product",
    "mall.write_product",
])

def main():
    # 1) 인증 URL 열기
    auth_url = (
        f"https://{MALL_ID}.cafe24api.com/api/v2/oauth/authorize"
        f"?response_type=code"
        f"&client_id={CLIENT_ID}"
        f"&state=shipment_manager"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
        f"&scope={SCOPES}"
    )

    print("=" * 60)
    print("브라우저에서 Cafe24 인증 페이지가 열립니다.")
    print("인증 후 리다이렉트된 URL 전체를 복사하세요.")
    print("=" * 60)
    print(f"\n인증 URL:\n{auth_url}\n")

    webbrowser.open(auth_url)

    # 2) 리다이렉트된 URL 입력받기
    print("-" * 60)
    redirected_url = input("리다이렉트된 URL을 붙여넣으세요: ").strip()

    # code 추출
    parsed = urllib.parse.urlparse(redirected_url)
    params = urllib.parse.parse_qs(parsed.query)

    if "code" not in params:
        # URL이 아니라 code만 붙여넣은 경우
        code = redirected_url
    else:
        code = params["code"][0]

    print(f"\n인증 코드: {code}")

    # 3) 토큰 교환
    token_url = f"https://{MALL_ID}.cafe24api.com/api/v2/oauth/token"
    auth_header = base64.b64encode(f"{CLIENT_ID}:{CLIENT_SECRET}".encode()).decode()

    resp = requests.post(token_url, headers={
        "Authorization": f"Basic {auth_header}",
        "Content-Type": "application/x-www-form-urlencoded",
    }, data={
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": REDIRECT_URI,
    })

    print(f"\n응답 코드: {resp.status_code}")
    result = resp.json()
    print(json.dumps(result, indent=2, ensure_ascii=False))

    if "access_token" in result:
        access_token = result["access_token"]
        refresh_token = result.get("refresh_token", "")

        # appsettings.json 자동 업데이트
        try:
            with open("appsettings.json", "r", encoding="utf-8") as f:
                config = json.load(f)
            config["Cafe24"]["AccessToken"] = access_token
            config["Cafe24"]["RefreshToken"] = refresh_token
            with open("appsettings.json", "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            print(f"\n✅ appsettings.json 업데이트 완료!")
        except Exception as e:
            print(f"\nappsettings.json 업데이트 실패: {e}")

        print(f"\n{'='*60}")
        print(f"ACCESS_TOKEN  = {access_token}")
        print(f"REFRESH_TOKEN = {refresh_token}")
        print(f"{'='*60}")
    else:
        print("\n❌ 토큰 발급 실패. 위 응답을 확인하세요.")

if __name__ == "__main__":
    main()
