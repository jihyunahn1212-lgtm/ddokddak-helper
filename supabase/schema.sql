-- ============================================================
-- 똑딱 정리 도우미: 작업 저장 테이블 + 접근 제한(RLS)
-- Supabase 대시보드 > SQL Editor > New query 에 붙여넣고 Run (한 번만 실행하면 돼요)
-- 여러 번 실행해도 안전합니다.
-- ============================================================

-- 1) 테이블: 사용자 한 명당, 파일 하나(fp)당 저장본 한 줄
create table if not exists public.work_saves (
  user_id  uuid        not null default auth.uid()
                       references auth.users (id) on delete cascade,
  -- 파일 구분용 SHA-256 해시(64자). 파일 이름·합계 같은 원본 정보는 저장하지 않아요.
  fp       text        not null check (char_length(fp) = 64),
  -- 작업 내용(일자 수정·제외한 행·추가 항목·체크·입력값). 너무 큰 데이터는 막아요.
  data     jsonb       not null check (octet_length(data::text) < 200000),
  saved_at timestamptz not null default now(),
  primary key (user_id, fp)
);

-- 2) 접근 제한(Row Level Security): 이게 꺼져 있으면 사이트 주소를 아는 누구나 읽을 수 있어요
alter table public.work_saves enable row level security;

-- 로그인하지 않은 사용자(anon)는 이 테이블에 아무것도 못 하게 하고,
-- 로그인한 사용자(authenticated)에게만 필요한 권한을 줘요.
revoke all on public.work_saves from anon;
grant select, insert, update, delete on public.work_saves to authenticated;

-- 3) 정책: 로그인한 사용자는 '자기 행'만 읽고 쓸 수 있어요
drop policy if exists "own rows: select" on public.work_saves;
drop policy if exists "own rows: insert" on public.work_saves;
drop policy if exists "own rows: update" on public.work_saves;
drop policy if exists "own rows: delete" on public.work_saves;

create policy "own rows: select" on public.work_saves
  for select to authenticated
  using ((select auth.uid()) = user_id);

create policy "own rows: insert" on public.work_saves
  for insert to authenticated
  with check ((select auth.uid()) = user_id);

create policy "own rows: update" on public.work_saves
  for update to authenticated
  using ((select auth.uid()) = user_id)
  with check ((select auth.uid()) = user_id);

create policy "own rows: delete" on public.work_saves
  for delete to authenticated
  using ((select auth.uid()) = user_id);
