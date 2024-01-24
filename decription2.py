from bs4 import BeautifulSoup

# 여기에 HTML 코드를 넣어주세요
html_content = """
<div class="memdoc">
<p>(description)이 함수는 두 정수를 입력으로 받아서 그 합을 반환합니다. 예를 들어, add(3, 4)를 호출하면 7이 반환됩니다.</p>
<dl class="params"><dt>Parameters</dt><dd>
  <table class="params">
    <tbody><tr><td class="paramdir">[in]</td><td class="paramname">in1_int</td><td>첫 번째 input </td></tr>
    <tr><td class="paramdir">[in]</td><td class="paramname">in2_double</td><td>두 번째 input </td></tr>
  </tbody></table>
  </dd>
</dl>
<dl class="section return"><dt>Returns</dt><dd>test return 0 </dd></dl>
</div>
"""

# BeautifulSoup을 사용하여 HTML 파싱
soup = BeautifulSoup(html_content, 'html.parser')

# <p> 태그 내용 추출
description = soup.find('div', class_='memdoc').find('p').get_text(strip=True)

# 파라미터 추출
params_dl = soup.find('div', class_='memdoc').find('dl', class_='params')
param_values = []
if params_dl:
    # 각 <tr>을 찾아서 <td> 값들을 가져옵니다.
    for tr in params_dl.find_all('tr'):
        # tr에서 <td> 태그들을 가져와서 파라미터 값으로 저장합니다.
        params = [td.get_text(strip=True) for td in tr.find_all('td') if not td.get('class')]
        param_values.extend(params)

# Return 값 추출
return_dd = soup.find('div', class_='memdoc').find('dl', class_='section return').find('dd').get_text(strip=True)

# 결과 리스트 구성
result_list = [description, param_values, return_dd]

# 결과 출력
print(result_list)
